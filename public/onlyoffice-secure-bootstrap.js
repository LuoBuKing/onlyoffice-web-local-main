/** OnlyOffice web-apps 在非 HTTPS 下 crypto.subtle 不可用，生成 __guid 时会抛 digest 异常 */
;(function () {
  try {
    if (!localStorage.getItem('__guid')) {
      localStorage.setItem(
        '__guid',
        Date.now() + '-' + Math.random().toString(36).slice(2, 11),
      )
    }
  } catch (e) {
    /* 隐私模式等无法写 localStorage 时忽略 */
  }

  var root = typeof globalThis !== 'undefined' ? globalThis : window
  if (!root.crypto) root.crypto = {}
  if (!root.crypto.subtle) {
    root.crypto.subtle = {
      digest: function (_algo, data) {
        var bytes =
          data instanceof ArrayBuffer
            ? new Uint8Array(data)
            : new Uint8Array(data.buffer || data)
        var out = new Uint8Array(20)
        for (var i = 0; i < bytes.length; i++) out[i % 20] ^= bytes[i]
        return Promise.resolve(out.buffer)
      },
    }
  }
})()
