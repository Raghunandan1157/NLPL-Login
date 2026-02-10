(function () {
  'use strict';

  var employeeData = {};
  var selectedId = null;
  var dataLoaded = false;

  // --- DOM elements ---
  var triggerInput = document.getElementById('triggerInput');
  var searchTrigger = document.getElementById('searchTrigger');
  var selectedCard = document.getElementById('selectedCard');
  var selectedAvatar = document.getElementById('selectedAvatar');
  var selectedName = document.getElementById('selectedName');
  var selectedIdText = document.getElementById('selectedIdText');
  var changeBtn = document.getElementById('changeBtn');
  var proceedBtn = document.getElementById('proceedBtn');
  var searchOverlay = document.getElementById('searchOverlay');
  var nameInput = document.getElementById('nameInput');
  var resultsList = document.getElementById('resultsList');
  var backBtn = document.getElementById('backBtn');
  var successOverlay = document.getElementById('successOverlay');
  var successName = document.getElementById('successName');
  var successId = document.getElementById('successId');
  var successBranch = document.getElementById('successBranch');
  var cardEl = document.querySelector('.card');

  var avatarColors = [
    ['#0E8B7D','#E0F5F2'], ['#6C5CE7','#EDE9FF'], ['#E17055','#FFEEE8'],
    ['#00B894','#E0FFF5'], ['#FDCB6E','#FFF8E7'], ['#E84393','#FFE4F3']
  ];

  // --- Show loading skeleton ---
  function showSkeleton() {
    cardEl.querySelector('.card-header').style.display = 'none';
    searchTrigger.style.display = 'none';
    proceedBtn.style.display = 'none';
    var skel = document.createElement('div');
    skel.id = 'skeleton';
    skel.innerHTML =
      '<div class="skeleton-field"><div class="skeleton-label"></div><div class="skeleton-input"></div></div>' +
      '<div class="skeleton-btn"></div>';
    cardEl.appendChild(skel);
  }

  function hideSkeleton() {
    var skel = document.getElementById('skeleton');
    if (skel) skel.remove();
    cardEl.querySelector('.card-header').style.display = '';
    searchTrigger.style.display = '';
    proceedBtn.style.display = '';
  }

  // --- Load data ---
  showSkeleton();
  fetch('employees.json')
    .then(function (r) { if (!r.ok) throw new Error('fail'); return r.json(); })
    .then(function (d) {
      employeeData = d;
      dataLoaded = true;
      hideSkeleton();
    })
    .catch(function (e) {
      console.error('Load error:', e);
      hideSkeleton();
    });

  // --- Helpers ---
  function getInitials(name) {
    var parts = name.trim().split(/\s+/);
    if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    return name.slice(0, 2).toUpperCase();
  }

  function search(query) {
    if (!query || query.length < 1) return [];
    var q = query.toLowerCase();
    var results = [];
    var keys = Object.keys(employeeData);
    for (var i = 0; i < keys.length; i++) {
      var id = keys[i];
      var emp = employeeData[id];
      var name = (typeof emp === 'string') ? emp : (emp.name || '');
      var branch = (typeof emp === 'string') ? '' : (emp.branch || '');
      var role = (typeof emp === 'string') ? '' : (emp.role || '');
      var designation = (typeof emp === 'string') ? '' : (emp.designation || '');

      var matchedField = '';
      if (name.toLowerCase().indexOf(q) !== -1) matchedField = 'name';
      else if (id.toLowerCase().indexOf(q) !== -1) matchedField = 'id';
      else if (branch.toLowerCase().indexOf(q) !== -1) matchedField = 'branch';
      else if (role.toLowerCase().indexOf(q) !== -1) matchedField = 'role';
      else if (designation.toLowerCase().indexOf(q) !== -1) matchedField = 'designation';

      if (matchedField) {
        results.push({
          id: id, name: name, branch: branch, role: role,
          designation: designation, matchedField: matchedField
        });
      }
      if (results.length >= 15) break;
    }
    return results;
  }

  function highlight(text, query) {
    var idx = text.toLowerCase().indexOf(query.toLowerCase());
    if (idx === -1) return text;
    return text.slice(0, idx) + '<span class="hl">' + text.slice(idx, idx + query.length) + '</span>' + text.slice(idx + query.length);
  }

  function esc(str) {
    return str.replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // --- Recent searches (localStorage) ---
  var RECENT_KEY = 'nlpl_recent';

  function getRecent() {
    try {
      var data = JSON.parse(localStorage.getItem(RECENT_KEY));
      return Array.isArray(data) ? data.slice(0, 5) : [];
    } catch (e) { return []; }
  }

  function saveRecent(emp) {
    var recent = getRecent().filter(function (r) { return r.id !== emp.id; });
    recent.unshift({ id: emp.id, name: emp.name, branch: emp.branch || '' });
    if (recent.length > 5) recent = recent.slice(0, 5);
    localStorage.setItem(RECENT_KEY, JSON.stringify(recent));
  }

  function renderRecent() {
    var recent = getRecent();
    resultsList.innerHTML = '';

    if (recent.length > 0) {
      var header = document.createElement('li');
      header.className = 'recent-header';
      header.innerHTML =
        '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>' +
        'Recent';
      resultsList.appendChild(header);

      recent.forEach(function (r, i) {
        var colors = avatarColors[i % avatarColors.length];
        var li = document.createElement('li');
        li.style.animationDelay = (i * 30) + 'ms';
        li.innerHTML =
          '<div class="avatar" style="color:' + colors[0] + ';background:' + colors[1] + '">' + getInitials(r.name) + '</div>' +
          '<div class="info">' +
            '<span class="name">' + esc(r.name) + '</span>' +
            (r.branch ? '<span class="branch">' + esc(r.branch) + '</span>' : '') +
          '</div>' +
          '<span class="emp-id">' + r.id + '</span>';
        li.addEventListener('click', function () {
          this.style.background = '#E0F5F2';
          this.style.transform = 'scale(0.98)';
          var emp = { id: r.id, name: r.name, branch: r.branch };
          setTimeout(function () { selectEmployee(emp); }, 150);
        });
        resultsList.appendChild(li);
      });
    } else {
      showPrompt();
    }
  }

  // --- Open/close search ---
  function openSearch() {
    if (!dataLoaded) return;
    searchOverlay.classList.add('active');
    nameInput.value = '';
    renderRecent();
    setTimeout(function () { nameInput.focus(); }, 200);
  }

  function closeSearch() {
    searchOverlay.classList.remove('active');
    nameInput.blur();
  }

  function showPrompt() {
    resultsList.innerHTML =
      '<li class="search-prompt">' +
        '<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>' +
        '<span>Search by name, employee ID, branch, or designation</span>' +
      '</li>';
  }

  // --- Render results ---
  function renderResults(results, query) {
    resultsList.innerHTML = '';

    if (!results.length && query.length >= 2) {
      var noLi = document.createElement('li');
      noLi.className = 'no-results';
      noLi.innerHTML =
        '<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">' +
          '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>' +
        '</svg>' +
        '<span>No employee found for "<strong>' + esc(query) + '</strong>"</span>';
      resultsList.appendChild(noLi);
      return;
    }

    if (!results.length) {
      renderRecent();
      return;
    }

    results.forEach(function (r, i) {
      var colors = avatarColors[i % avatarColors.length];
      var li = document.createElement('li');
      li.style.animationDelay = (i * 30) + 'ms';

      // Build match tag for non-name matches
      var matchTag = '';
      if (r.matchedField === 'id') {
        matchTag = '<span class="match-tag">ID: ' + highlight(r.id, query) + '</span>';
      } else if (r.matchedField === 'branch') {
        matchTag = '<span class="match-tag">Branch: ' + highlight(r.branch, query) + '</span>';
      } else if (r.matchedField === 'role') {
        matchTag = '<span class="match-tag">Role: ' + highlight(r.role, query) + '</span>';
      } else if (r.matchedField === 'designation') {
        matchTag = '<span class="match-tag">Designation: ' + highlight(r.designation, query) + '</span>';
      }

      var nameHtml = r.matchedField === 'name' ? highlight(r.name, query) : esc(r.name);

      li.innerHTML =
        '<div class="avatar" style="color:' + colors[0] + ';background:' + colors[1] + '">' + getInitials(r.name) + '</div>' +
        '<div class="info">' +
          '<span class="name">' + nameHtml + '</span>' +
          (r.branch && r.matchedField !== 'branch' ? '<span class="branch">' + esc(r.branch) + '</span>' : '') +
          matchTag +
        '</div>' +
        '<span class="emp-id">' + r.id + '</span>';

      li.addEventListener('click', function () {
        this.style.background = '#E0F5F2';
        this.style.transform = 'scale(0.98)';
        var emp = r;
        setTimeout(function () { selectEmployee(emp); }, 150);
      });
      resultsList.appendChild(li);
    });
  }

  // --- Select employee ---
  function selectEmployee(emp) {
    selectedId = emp.id;
    saveRecent(emp);
    closeSearch();

    searchTrigger.style.display = 'none';
    selectedCard.classList.add('visible');
    selectedAvatar.textContent = getInitials(emp.name);
    selectedName.textContent = emp.name;
    selectedIdText.textContent = emp.id;

    proceedBtn.disabled = false;

    sessionStorage.setItem('employeeId', emp.id);
    sessionStorage.setItem('employeeName', emp.name);
    sessionStorage.setItem('employeeBranch', emp.branch || '');
  }

  // --- Clear selection ---
  function clearSelection() {
    selectedId = null;
    selectedCard.classList.remove('visible');
    searchTrigger.style.display = 'block';
    proceedBtn.disabled = true;
  }

  // --- Events ---
  triggerInput.addEventListener('click', openSearch);
  triggerInput.addEventListener('focus', function () { this.blur(); openSearch(); });
  backBtn.addEventListener('click', closeSearch);

  nameInput.addEventListener('input', function () {
    var query = this.value.trim();
    renderResults(search(query), query);
  });

  changeBtn.addEventListener('click', clearSelection);

  // Proceed → success screen
  proceedBtn.addEventListener('click', function () {
    if (!selectedId) return;
    var emp = employeeData[selectedId];
    var name = (typeof emp === 'string') ? emp : emp.name;
    var branch = (typeof emp === 'string') ? '' : (emp.branch || '');

    successName.textContent = 'Welcome, ' + name + '!';
    successId.textContent = selectedId;
    successBranch.textContent = branch;

    proceedBtn.style.transform = 'scale(0.95)';
    proceedBtn.disabled = true;

    setTimeout(function () {
      successOverlay.classList.add('show');
      // Redirect to dashboard after success animation
      setTimeout(function () {
        window.location.href = 'dashboard.html';
      }, 1500);
    }, 300);
  });
})();
