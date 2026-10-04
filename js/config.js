/* ================================================================
   Kuwait Hospital – Meeting Room Booking
   Apps Script backend connection
   IT Department – Mohamed Sirajudeen
   ================================================================ */

window.KH_API_URL = 'https://script.google.com/macros/s/AKfycbwM9MzgGc_8xWvrqgB80MDjIGNCyrbj15OHZBYXPY1bEpLURpBDZOrqdfV-j-RESyBg/exec';

(function () {
  function call(fn, args, ok, fail) {
    if (!window.KH_API_URL || window.KH_API_URL.indexOf('/exec') < 0) {
      if (fail) fail(new Error('Setup needed: put the Apps Script /exec URL in js/config.js'));
      return;
    }
    fetch(window.KH_API_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body:    JSON.stringify({ fn: fn, args: args })
    })
    .then(function (r) { return r.json(); })
    .then(function (res) {
      if (res && res.ok) { if (ok) ok(res.data); }
      else if (fail) fail(new Error((res && res.error) || 'Server error'));
    })
    .catch(function () {
      if (fail) fail(new Error('Cannot reach the server. Check connection and try again.'));
    });
  }

  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () { call(name, Array.prototype.slice.call(arguments), ok, fail); };
      }
    });
  }

  window.google = { script: { get run() { return runner(null, null); } } };
})();
