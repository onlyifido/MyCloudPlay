"""Browser regressions. Remote identity, OneDrive, and audio use deterministic fixtures."""
import functools
import http.server
import io
import json
import os
from pathlib import Path
import shutil
import threading
import unittest
import wave

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
MSAL_FIXTURE = """
window.msal = {
  InteractionRequiredAuthError: class extends Error {},
  PublicClientApplication: class {
    handleRedirectPromise() { return Promise.resolve(null); }
    getAllAccounts() { return []; }
    getAccountByHomeId() { return null; }
    loginRedirect() { window.testLoginCalls = (window.testLoginCalls || 0) + 1; return Promise.resolve(); }
    logoutRedirect() { return Promise.resolve(); }
  }
};
"""

def audio_fixture():
    data = io.BytesIO()
    with wave.open(data, 'wb') as audio:
        audio.setnchannels(1)
        audio.setsampwidth(2)
        audio.setframerate(8000)
        audio.writeframes(b'\0\0' * 8000 * 30)
    return data.getvalue()


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass


class BrowserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = http.server.ThreadingHTTPServer(('127.0.0.1', 0), functools.partial(QuietHandler, directory=str(ROOT)))
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.origin = f'http://127.0.0.1:{cls.server.server_port}'
        cls.playwright = sync_playwright().start()
        executable = os.environ.get('CHROMIUM_EXECUTABLE') or shutil.which('chromium')
        cls.browser = cls.playwright.chromium.launch(executable_path=executable, args=['--no-sandbox'])
        cls.audio = audio_fixture()

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.playwright.stop()
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()

    def setUp(self):
        self.context = self.browser.new_context(viewport={'width': 1440, 'height': 900})
        self.context.route('https://**/*', self.external)
        self.page = self.context.new_page()
        self.errors = []
        self.page.on('pageerror', lambda error: self.errors.append(str(error)))

    def tearDown(self):
        self.context.close()
        self.assertEqual(self.errors, [], 'Unexpected browser errors')

    def external(self, route):
        url = route.request.url
        if 'alcdn.msauth.net/' in url:
            route.fulfill(body=MSAL_FIXTURE, content_type='application/javascript')
        elif 'soundhelix.com/' in url or 'media.example.test/' in url:
            route.fulfill(body=self.audio, content_type='audio/wav')
        else:
            route.abort()

    def load(self):
        self.page.goto(self.origin, wait_until='load')
        self.page.wait_for_selector('#track_row_sample_1')

    def json_response(self, route, body, status=200):
        route.fulfill(status=status, body=json.dumps(body), content_type='application/json')

    def use_test_account(self):
        self.page.evaluate("() => { getToken = async () => 'test-token'; }")

    def test_appearance_persists_and_follows_system(self):
        self.load()
        self.page.locator('#appearance_button').click()
        colors = set()
        for skin in ['ocean', 'moss', 'sand', 'rose', 'graphite']:
            self.page.locator(f'button[data-skin="{skin}"]').click()
            for mode in ['light', 'dark']:
                self.page.locator(f'[data-mode="{mode}"]').click()
                colors.add(self.page.evaluate("getComputedStyle(document.documentElement).getPropertyValue('--accent').trim()"))
                self.assertEqual(self.page.locator(f'button[data-skin="{skin}"]').get_attribute('aria-pressed'), 'true')
                canvas = self.page.evaluate("getComputedStyle(document.documentElement).getPropertyValue('--canvas').trim()")
                expected = 'rgb(' + ', '.join(str(int(canvas[i:i+2], 16)) for i in (1, 3, 5)) + ')'
                self.page.wait_for_function('(color) => getComputedStyle(document.body).backgroundColor === color', arg=expected)
        self.assertEqual(len(colors), 10)
        self.page.keyboard.press('Escape')
        self.assertTrue(self.page.locator('#appearance_button').evaluate('(e) => e === document.activeElement'))
        self.page.reload(wait_until='load')
        self.assertEqual(self.page.evaluate("[appearanceSkin, appearanceMode, document.documentElement.classList.contains('dark')]"), ['graphite', 'dark', True])
        self.page.evaluate("chooseMode('system')")
        self.page.emulate_media(color_scheme='light')
        self.assertFalse(self.page.evaluate("document.documentElement.classList.contains('dark')"))
        self.page.emulate_media(color_scheme='dark')
        self.page.wait_for_function("document.documentElement.classList.contains('dark')")
        self.assertTrue(self.page.evaluate("document.documentElement.classList.contains('dark')"))

    def test_responsive_controls_do_not_overlap(self):
        self.load()
        for width, height in [(320, 568), (360, 640), (390, 844), (768, 1024), (1440, 900), (667, 375)]:
            self.page.set_viewport_size({'width': width, 'height': height})
            boxes = self.page.locator('.player-controls button').evaluate_all('(els) => els.map(e => { const r=e.getBoundingClientRect(); return {left:r.left,right:r.right,top:r.top,bottom:r.bottom,width:r.width,height:r.height}; })')
            for index, box in enumerate(boxes):
                self.assertGreaterEqual(box['width'], 44)
                self.assertGreaterEqual(box['height'], 44)
                self.assertGreaterEqual(box['left'], 0)
                self.assertLessEqual(box['right'], width)
                if index:
                    self.assertLessEqual(boxes[index-1]['right'], box['left'])
            self.assertTrue(self.page.locator('#login_button').evaluate('(e) => e.getBoundingClientRect().right <= innerWidth'))
            self.assertTrue(self.page.locator('#appearance_button').evaluate('(e) => e.getBoundingClientRect().right <= innerWidth'))
            self.assertTrue(self.page.locator('#track_row_sample_1').is_visible())

    def test_modal_focus_keyboard_and_english_labels(self):
        self.load()
        self.page.locator('#appearance_button').focus()
        self.page.keyboard.press('Enter')
        self.assertTrue(self.page.locator('#appearance_modal').is_visible())
        self.page.locator('button[data-skin="graphite"]').focus()
        self.page.keyboard.press('Tab')
        self.assertTrue(self.page.get_by_role('button', name='Close appearance').evaluate('(e) => e === document.activeElement'))
        self.page.keyboard.press('Escape')
        self.assertFalse(self.page.locator('#appearance_modal').is_visible())
        self.page.get_by_role('button', name='About', exact=True).click()
        self.assertIn('Google Analytics', self.page.locator('#info_modal').inner_text())
        self.page.keyboard.press('Escape')
        self.assertFalse(self.page.locator('#info_modal').is_visible())
        unlabeled = self.page.locator('button:visible').evaluate_all('(els) => els.filter(e => !e.textContent.trim() && !e.getAttribute("aria-label") && !e.title).length')
        self.assertEqual(unlabeled, 0)
        self.assertNotRegex(self.page.locator('body').inner_text(), '[가-힣]')
        self.assertEqual(self.page.locator('html').get_attribute('lang'), 'en')

    def test_corrupt_and_disabled_storage_do_not_block_startup(self):
        self.page.add_init_script("localStorage.setItem('folder_cache', 'bad-json'); localStorage.setItem('recent_tracks', '{}'); localStorage.setItem('repeatMode', 'bad'); localStorage.setItem('skin', 'invalid');")
        self.load()
        self.assertEqual(self.page.evaluate('[repeatMode, appearanceSkin, readRecentTracks().length]'), [0, 'ocean', 0])
        self.page.add_init_script("Storage.prototype.getItem = () => { throw new DOMException('Blocked'); }; Storage.prototype.setItem = () => { throw new DOMException('Blocked'); };")
        self.page.reload(wait_until='load')
        self.assertEqual(self.page.locator('[id^="track_row_"]').count(), 3)
        self.page.evaluate("chooseMode('dark'); chooseSkin('moss')")
        self.assertEqual(self.page.evaluate('appearanceSkin'), 'moss')

    def test_repeat_shuffle_and_sample_history(self):
        self.load()
        self.page.locator('#repeat_btn').click()
        self.page.locator('#repeat_btn').click()
        self.page.locator('#shuffle_btn').click()
        self.page.reload(wait_until='load')
        self.assertEqual(self.page.evaluate('[repeatMode, audio.loop, isShuffle]'), [2, True, True])
        self.page.evaluate("() => { window.testTokenCalls = 0; getToken = async () => { window.testTokenCalls++; return null; }; }")
        self.page.locator('#track_row_sample_1 button').click()
        self.page.wait_for_function('!audio.paused && audio.currentTime > 0')
        self.assertEqual(self.page.locator('#track_icon_sample_1 .eq-bar').count(), 3)
        self.assertEqual(self.page.locator('#play_pause_btn').get_attribute('aria-label'), 'Pause')
        self.page.locator('#play_pause_btn').click()
        self.assertTrue(self.page.evaluate('audio.paused'))
        self.page.evaluate('playRecentTrack({...sampleTracks[1], folderId: "root", folderName: "My Drive"})')
        self.page.wait_for_function('currentIndex === 1 && !audio.paused')
        self.assertEqual(self.page.evaluate('window.testTokenCalls'), 0)
        self.assertEqual(self.page.locator('#file_info').inner_text(), 'Sample Track 2 - Upbeat')

    def test_external_metadata_is_text(self):
        self.load()
        self.page.evaluate('''() => {
            playingQueue = [{id:'test', name:'<img id="injected-name" src=x>.mp3', audio:{title:'<b id="injected-title">Title</b>',artist:'<img id="injected-artist">'}}];
            currentIndex = 0; showTrackInfo();
        }''')
        self.assertEqual(self.page.locator('#injected-name, #injected-title, #injected-artist').count(), 0)
        self.assertIn('<b id="injected-title">Title</b>', self.page.locator('#track_info_content').inner_text())

    def test_paginated_files_remain_sorted_by_name(self):
        calls = []
        def graph(route):
            calls.append(route.request.url)
            if 'page=2' in route.request.url:
                self.json_response(route, {'value': [{'id':'a', 'name':'A.mp3', 'file':{'mimeType':'audio/mpeg'}}]})
            else:
                self.json_response(route, {'value': [{'id':'z', 'name':'Z.mp3', 'file':{'mimeType':'audio/mpeg'}}], '@odata.nextLink':'https://graph.microsoft.com/v1.0/me/drive/items/root/children?page=2'})
        self.context.route('https://graph.microsoft.com/**', graph)
        self.load()
        self.use_test_account()
        self.page.evaluate("fetchMyFolders('root', 'My Drive')")
        self.assertEqual(len(calls), 2)
        self.assertEqual(self.page.locator('[id^="track_text_"]').all_text_contents(), ['A', 'Z'])
        self.assertEqual(self.page.evaluate('currentFolderAudios.map(t => t.folderId)'), ['root', 'root'])

    def test_stale_track_request_cannot_replace_new_selection(self):
        self.load()
        result = self.page.evaluate('''async () => {
            let releaseFirst;
            graphRequest = path => path.includes('/first') ? new Promise(resolve => releaseFirst=resolve) : Promise.resolve({'@microsoft.graph.downloadUrl':'https://media.example.test/second.wav'});
            playingQueue = [{id:'first',name:'First.mp3'},{id:'second',name:'Second.mp3'}];
            const first = playTrack(0, false);
            await playTrack(1, false);
            releaseFirst({'@microsoft.graph.downloadUrl':'https://media.example.test/first.wav'});
            await first;
            return {index:currentIndex,src:audio.src,label:fileInfoEl.textContent};
        }''')
        self.assertEqual(result['index'], 1)
        self.assertEqual(result['label'], 'Second')
        self.assertTrue(result['src'].endswith('second.wav'))

    def test_failed_rename_keeps_dialog_and_original_name(self):
        self.context.route('https://graph.microsoft.com/**', lambda route: self.json_response(route, {'error':{}}, 403))
        self.load()
        self.use_test_account()
        self.page.evaluate("requestRename('test', 'Original.mp3', false)")
        self.page.locator('#rename_textarea').fill('Changed')
        self.page.locator('#rename_confirm_btn').click()
        self.page.wait_for_function("!document.getElementById('rename_confirm_btn').disabled")
        self.assertTrue(self.page.locator('#rename_modal').is_visible())
        self.assertIn('did not allow', self.page.locator('#notice_text').inner_text())

    def test_upload_uses_fragments_and_reports_success(self):
        ranges = []
        def graph(route):
            if route.request.method == 'POST':
                self.json_response(route, {'uploadUrl':'https://upload.example.test/session'})
            else:
                self.json_response(route, {'value':[]})
        def upload(route):
            value = route.request.headers['content-range']
            ranges.append(value)
            final = value.startswith('bytes 3276800-')
            self.json_response(route, {'id':'uploaded'} if final else {'nextExpectedRanges':['3276800-']}, 201 if final else 202)
        self.context.route('https://graph.microsoft.com/**', graph)
        self.context.route('https://upload.example.test/**', upload)
        self.load()
        self.use_test_account()
        self.page.evaluate("currentFolderId='root'; pathHistory=[{id:'root',name:'My Drive'}]")
        self.page.locator('#file_upload').set_input_files({'name':'Recording.mp3','mimeType':'audio/mpeg','buffer':b'a' * (3276800 + 1024)})
        self.page.wait_for_function('uploadController === null')
        self.assertEqual(ranges, ['bytes 0-3276799/3277824', 'bytes 3276800-3277823/3277824'])
        self.assertEqual(self.page.locator('#notice_text').inner_text(), '1 file uploaded.')

    def test_end_of_queue_stays_finished(self):
        self.load()
        self.page.locator('#track_row_sample_3 button').click()
        self.page.wait_for_function('!audio.paused && audio.currentTime > 0')
        self.page.get_by_role('button', name='Next track', exact=True).click()
        self.page.wait_for_function('audio.readyState === 0 && audio.paused')
        self.assertEqual(self.page.locator('#status_info').inner_text(), 'FINISHED')
        self.assertEqual(self.page.locator('#play_pause_btn').get_attribute('aria-label'), 'Play')
        self.assertTrue(self.page.locator('#progress_bar').is_disabled())

    def test_upload_cancellation_cleans_up_session(self):
        methods = []
        pending = []
        def graph(route):
            self.json_response(route, {'uploadUrl':'https://upload.example.test/cancel'} if route.request.method == 'POST' else {'value':[]})
        def upload(route):
            methods.append(route.request.method)
            if route.request.method == 'DELETE':
                route.fulfill(status=204, body='')
            else:
                pending.append(route)
            # Leave PUT pending until the Cancel button aborts its request.
        self.context.route('https://graph.microsoft.com/**', graph)
        self.context.route('https://upload.example.test/**', upload)
        self.load()
        self.use_test_account()
        self.page.evaluate("currentFolderId='root'; pathHistory=[{id:'root',name:'My Drive'}]")
        with self.page.expect_request('https://upload.example.test/cancel'):
            self.page.locator('#file_upload').set_input_files({'name':'Cancel.mp3','mimeType':'audio/mpeg','buffer':b'a' * 1024})
        self.page.get_by_role('button', name='Cancel upload', exact=True).click()
        self.page.wait_for_function('uploadController === null')
        self.assertEqual(methods, ['PUT', 'DELETE'])
        self.assertEqual(self.page.locator('#notice_text').inner_text(), 'Upload cancelled. 0 completed.')
        self.assertEqual(self.page.locator('#file_upload').input_value(), '')
        for route in pending:
            route.abort()

    def test_failed_delete_shows_error_without_refreshing(self):
        calls = []
        def denied(route):
            calls.append(route.request.method)
            self.json_response(route, {'error':{}}, 403)
        self.context.route('https://graph.microsoft.com/**', denied)
        self.load()
        self.use_test_account()
        self.page.on('dialog', lambda dialog: dialog.accept())
        self.page.evaluate("deleteItem('test', 'Track.mp3', false)")
        self.assertEqual(calls, ['DELETE'])
        self.assertIn('did not allow', self.page.locator('#notice_text').inner_text())


if __name__ == '__main__':
    unittest.main()
