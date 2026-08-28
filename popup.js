const btn = document.getElementById('authorize');
const status = document.getElementById('status');

btn.addEventListener('click', async () => {
  btn.disabled = true;
  status.textContent = '承認画面を開いています…';
  try {
    const res = await chrome.runtime.sendMessage({type: 'JEFF_AUTHORIZE'});
    if (!res?.ok) throw new Error(res?.error || 'Authorization failed');
    status.textContent = '承認完了。Google Sheetsでショートカットを使用できます。';
    status.className = 'ok';
  } catch (e) {
    status.textContent = 'エラー: ' + (e?.message || String(e));
  } finally {
    btn.disabled = false;
  }
});
