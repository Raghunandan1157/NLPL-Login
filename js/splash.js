(function () {
  // Only play once per session
  if (sessionStorage.getItem('splashShown') === 'true') {
    var el = document.getElementById('splashOverlay');
    if (el) el.style.display = 'none';
    return;
  }

  var terminal = document.getElementById('splashTerminal');
  var overlay = document.getElementById('splashOverlay');
  if (!terminal || !overlay) return;

  function pad(n) { return n < 10 ? '0' + n : '' + n; }

  function ts(d) {
    return pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  }

  var clock = new Date();

  function tick() {
    clock = new Date(clock.getTime() + (Math.floor(Math.random() * 2000) + 1000));
    return ts(clock);
  }

  function addLine(timestamp, text) {
    var line = document.createElement('div');
    line.className = 'splash-line';
    line.innerHTML =
      '<span class="splash-ts">' + timestamp + '</span>' +
      '<span class="splash-msg">' + text + '</span>';
    terminal.appendChild(line);
    setTimeout(function () { line.classList.add('visible'); }, 10);
    return line;
  }

  function addAscii(timestamp) {
    var line = document.createElement('div');
    line.className = 'splash-line';
    var tsEl = document.createElement('span');
    tsEl.className = 'splash-ts';
    tsEl.textContent = timestamp;
    line.appendChild(tsEl);

    var art = document.createElement('pre');
    art.className = 'splash-ascii';
    art.textContent =
      '\n' +
      '  ╔══════════════════════════════════════════╗\n' +
      '  ║                                          ║\n' +
      '  ║   ███╗   ██╗ ██╗     ██████╗  ██╗       ║\n' +
      '  ║   ████╗  ██║ ██║     ██╔══██╗ ██║       ║\n' +
      '  ║   ██╔██╗ ██║ ██║     ██████╔╝ ██║       ║\n' +
      '  ║   ██║╚██╗██║ ██║     ██╔═══╝  ██║       ║\n' +
      '  ║   ██║ ╚████║ ███████╗██║      ███████╗  ║\n' +
      '  ║   ╚═╝  ╚═══╝ ╚══════╝╚═╝      ╚══════╝  ║\n' +
      '  ║                                          ║\n' +
      '  ╚══════════════════════════════════════════╝\n';
    line.appendChild(art);
    terminal.appendChild(line);
    setTimeout(function () { line.classList.add('visible'); }, 10);
  }

  var steps = [
    'INCOMING HTTP REQUEST DETECTED ...',
    'SERVICE WAKING UP ...',
    null, // ASCII art
    'LOADING EMPLOYEE DATABASE ...',
    'INITIALIZING FLASK SERVER ...',
    'CONNECTING TO APPLICATION ...',
    'COMPILING STATIC ASSETS ...',
    'ENVIRONMENT VARIABLES INJECTED ...',
    'RUNNING HEALTH CHECKS ...',
    'ALL SYSTEMS OPERATIONAL ...',
    'STEADY HANDS. YOUR APP IS LIVE ...'
  ];

  var i = 0;

  function next() {
    if (i >= steps.length) {
      setTimeout(function () {
        overlay.classList.add('done');
        sessionStorage.setItem('splashShown', 'true');
        setTimeout(function () { overlay.style.display = 'none'; }, 500);
      }, 800);
      return;
    }

    var t = tick();
    if (steps[i] === null) {
      addAscii(t);
    } else {
      addLine(t, steps[i]);
    }
    i++;
    setTimeout(next, Math.floor(Math.random() * 150) + 350);
  }

  next();
})();
