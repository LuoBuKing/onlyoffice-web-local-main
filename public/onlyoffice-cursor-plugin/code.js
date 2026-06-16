/* global Asc, Api, window */
;(function (window) {
  var MSG = 'ONLYOFFICE_CURSOR_PROBE'
  var US_NORMAL_TWIPS = 1440
  var tmr = null

  function post(payload) {
    var data = { type: MSG, payload: payload }
    try {
      if (window.top && window.top !== window) {
        window.top.postMessage(data, '*')
        return
      }
    } catch (e) {}
    if (window.parent && window.parent !== window) {
      window.parent.postMessage(data, '*')
    }
  }

  function probe(lastEvent, cb) {
    if (!window.Asc || !window.Asc.plugin || !window.Asc.plugin.info) {
      cb({ event: lastEvent, doc: { error: 'no plugin info' } })
      return
    }
    Asc.scope._ed = window.Asc.plugin.info.editorType
    window.Asc.plugin.callCommand(
      function () {
        var ed = Asc.scope._ed
        if (ed === 'cell') {
          try {
            var sh = Api.GetActiveSheet()
            var rng = sh.GetSelection()
            return {
              editor: 'cell',
              address: rng.GetAddress(),
              row: rng.GetRow(),
              col: rng.GetCol(),
              text: String(rng.GetText ? rng.GetText() : '').slice(0, 200),
            }
          } catch (e1) {
            return { editor: 'cell', error: String(e1) }
          }
        }
        if (ed === 'word') {
          try {
            var doc = Api.GetDocument()
            var out = { editor: 'word' }
            try {
              var cp = doc.GetCurrentParagraph && doc.GetCurrentParagraph()
              if (cp && cp.GetText) out.currentParagraphText = String(cp.GetText()).slice(0, 300)
            } catch (e2) {}
            try {
              var rs = doc.GetRangeBySelect && doc.GetRangeBySelect()
              if (rs && rs.GetText) out.selectionText = String(rs.GetText()).slice(0, 300)
            } catch (e3) {}
            return out
          } catch (e4) {
            return { editor: 'word', error: String(e4) }
          }
        }
        return { editor: ed || 'unknown' }
      },
      false,
      false,
      function (ret) {
        cb({ event: lastEvent, doc: ret })
      },
    )
  }

  function emit(lastEvent) {
    if (tmr) return
    tmr = setTimeout(function () {
      tmr = null
      probe(lastEvent, function (payload) {
        post(payload)
      })
    }, 50)
  }

  // 官方说明：init 与 button 均为必填接口，缺少 button 可能导致插件无法正常工作
  window.Asc.plugin.button = function () {}

  function applyUsNormalPageMargins(cb) {
    if (!window.Asc || !window.Asc.plugin) {
      if (cb) cb()
      return
    }
    var ed = window.Asc.plugin.info ? window.Asc.plugin.info.editorType : null
    if (ed !== 'word') {
      if (cb) cb()
      return
    }
    window.Asc.plugin.callCommand(
      function () {
        var section = Api.GetDocument().GetFinalSection()
        section.SetPageMargins(US_NORMAL_TWIPS, US_NORMAL_TWIPS, US_NORMAL_TWIPS, US_NORMAL_TWIPS)
        return true
      },
      false,
      false,
      function () {
        if (cb) cb()
      },
    )
  }

  window.Asc.plugin.init = function () {
    post({
      event: 'plugin_inited',
      editorType: window.Asc.plugin.info ? window.Asc.plugin.info.editorType : null,
    })

    function safeAttach(name, fn) {
      try {
        window.Asc.plugin.attachEditorEvent(name, fn)
      } catch (e) {
        post({ event: 'attachEditorEvent_failed', name: name, error: String(e) })
      }
    }

    safeAttach('onTargetPositionChanged', function () {
      emit('onTargetPositionChanged')
    })
    safeAttach('onSelectionEnd', function () {
      emit('onSelectionEnd')
    })
    safeAttach('onDocumentContentReady', function () {
      applyUsNormalPageMargins(function () {
        post({ event: 'onDocumentContentReady' })
      })
    })
  }
})(window)
