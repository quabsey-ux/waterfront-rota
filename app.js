// ═══════════════════════════════════════════════════════════════════
// WATERFRONT ROTA MANAGER — Application Logic
// ═══════════════════════════════════════════════════════════════════
(function () {
  'use strict';

  // ── CONFIGURATION ──────────────────────────────────────────────
  const API_URL = 'https://script.google.com/macros/s/AKfycbwVy1VzsPAaeEvhpLg6kbPWtyR055DTGtFVd0dfrpeaddjurdqxEC0hhpZZTLRga44m/exec';

  const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  const FULL_DAYS = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'];
  const ROOMS = ['ROOM 1', 'ROOM 2 - PRE-OP', 'ROOM 3', 'ROOM 4 - TRT RM', 'MINOR THEATRE', 'MAIN THEATRE'];
  const MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

  const SHIFT_PRESETS = [
    { label: 'Early Scrub', value: '07:00-15:00 S' },
    { label: 'Long Scrub', value: '07:00-19:30 S' },
    { label: 'Ward Long', value: '07:30-19:30 W' },
    { label: 'Reception', value: '09:30-19:30 R' },
    { label: 'OPD', value: '08:00-19:00 OPD' },
    { label: 'Circ/Recovery', value: '07:00-15:00 CR' },
    { label: 'Day Off', value: 'DO' },
    { label: 'Annual Leave', value: 'A/L' },
    { label: 'Study Leave', value: 'S/L' },
    { label: 'Sick', value: 'SICK' },
    { label: 'Compassionate', value: 'C/L' },
    { label: 'Available', value: 'AV' },
    { label: 'WFH', value: 'WFH' },
  ];

  const LEGEND_ITEMS = [
    { cls: 'time', label: '09:00-17:00', desc: 'Timed Shift' },
    { cls: 'do', label: 'DO', desc: 'Day Off' },
    { cls: 'al', label: 'A/L', desc: 'Annual Leave' },
    { cls: 'av', label: 'AV', desc: 'Available' },
    { cls: 'wfh', label: 'WFH', desc: 'Work From Home' },
    { cls: 'scrub', label: 'S', desc: 'Scrub' },
    { cls: 'assist', label: 'A', desc: 'Assist' },
    { cls: 'ward', label: 'W', desc: 'Ward' },
    { cls: 'rec', label: 'R', desc: 'Reception' },
    { cls: 'cr', label: 'CR', desc: 'Circ/Recovery' },
    { cls: 'opd', label: 'OPD', desc: 'Outpatients' },
    { cls: 'sl', label: 'S/L', desc: 'Study Leave' },
    { cls: 'sick', label: 'SICK', desc: 'Sick Leave' },
    { cls: 'cl', label: 'C/L', desc: 'Compassionate Leave' },
  ];

  // ── STATE ──────────────────────────────────────────────────────
  const state = {
    weekOffset: 0,
    staff: [],
    currentRota: {},
    currentTheatre: {},
    leaveRequests: [],
    leaveBalances: [],
    isPublished: false,
    editingCell: null,
    editingStaffId: null,
    settings: { manager_emails: [], auto_email: true },
    settingsLoaded: false,
    legendVisible: false,
    currentTab: 'rota',
    staffViewMode: false,
    staffViewName: null,
    cache: {},
    undoStack: [],
    staffSearch: '',
    monthlyView: false,
    monthlyData: {},  // weekKey → { shifts, published }
    monthlyWeekKeys: null,
    adminPin: null,
    pinVerified: false,
    hasPinConfigured: false,
  };

  // ── UTILITIES ──────────────────────────────────────────────────
  function $(id) { return document.getElementById(id); }

  function escapeHtml(str) {
    if (str == null) return '';
    var s = String(str);
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(s));
    return div.innerHTML;
  }

  function showToast(msg, type, undoCallback) {
    var container = $('toast-container');
    container.innerHTML = '';
    var toast = document.createElement('div');
    toast.className = 'toast ' + (type || 'info');

    var span = document.createElement('span');
    span.textContent = msg;
    toast.appendChild(span);

    if (undoCallback) {
      var btn = document.createElement('button');
      btn.className = 'undo-btn';
      btn.textContent = 'Undo';
      btn.onclick = function () {
        undoCallback();
        toast.remove();
      };
      toast.appendChild(btn);
    }

    container.appendChild(toast);
    setTimeout(function () { toast.remove(); }, undoCallback ? 6000 : 4000);
  }

  function hideLoading() {
    var el = $('loading-skeleton');
    if (el) el.style.display = 'none';
    var appEl = $('app');
    if (appEl) appEl.style.display = 'block';
  }

  // ── API ────────────────────────────────────────────────────────
  function api(action, params) {
    params = params || {};
    if (API_URL === 'YOUR_APPS_SCRIPT_URL') {
      $('config-banner').style.display = 'block';
      return Promise.resolve(null);
    }
    var url = new URL(API_URL);
    url.searchParams.set('action', action);
    Object.keys(params).forEach(function (k) {
      var v = params[k];
      url.searchParams.set(k, typeof v === 'object' ? JSON.stringify(v) : v);
    });
    return fetch(url.toString())
      .then(function (res) { return res.json(); })
      .catch(function (e) {
        showToast('Connection error: ' + e.message, 'error');
        return null;
      });
  }

  // ── PIN SECURITY ─────────────────────────────────────────────
  function apiWrite(action, params) {
    params = params || {};
    params.pin = state.adminPin || sessionStorage.getItem('wfr_admin_pin') || '';
    return api(action, params).then(function (result) {
      if (result && result.error === 'INVALID_PIN') {
        sessionStorage.removeItem('wfr_admin_pin');
        state.adminPin = null;
        state.pinVerified = false;
        showToast('Invalid PIN — please try again', 'error');
        return null;
      }
      return result;
    });
  }

  function requirePin(callback) {
    if (!state.hasPinConfigured) { callback(); return; }
    if (state.pinVerified && state.adminPin) { callback(); return; }
    var cached = sessionStorage.getItem('wfr_admin_pin');
    if (cached) { state.adminPin = cached; state.pinVerified = true; callback(); return; }
    var pin = prompt('Enter admin PIN to make changes:');
    if (!pin) return;
    api('verifyPin', { pin: pin }).then(function (result) {
      if (result && result.valid) {
        state.adminPin = pin;
        state.pinVerified = true;
        sessionStorage.setItem('wfr_admin_pin', pin);
        callback();
      } else {
        showToast('Invalid PIN', 'error');
      }
    });
  }

  function saveAdminPin() {
    var pin = $('admin-pin-input').value.trim();
    if (!pin) { showToast('Please enter a PIN', 'warning'); return; }
    requirePin(function () {
      apiWrite('saveConfig', { config: JSON.stringify({ admin_pin: pin, manager_emails: state.settings.manager_emails, auto_email: state.settings.auto_email }) }).then(function (result) {
        if (result && result.success) {
          state.hasPinConfigured = true;
          // Clear session so user must enter the NEW pin on next edit
          state.adminPin = null;
          state.pinVerified = false;
          sessionStorage.removeItem('wfr_admin_pin');
          $('admin-pin-input').value = '';
          showToast('PIN set! You\u2019ll be prompted on your next edit.', 'success');
        } else if (!result) {
          showToast('Failed to save PIN \u2014 check the console', 'error');
        }
      });
    });
  }

  // ── WEEK CALCULATIONS ─────────────────────────────────────────
  function getWeekDates() {
    var today = new Date();
    var monday = new Date(today);
    monday.setDate(today.getDate() - ((today.getDay() + 6) % 7) + state.weekOffset * 7);
    return DAYS.map(function (_, i) {
      var d = new Date(monday);
      d.setDate(monday.getDate() + i);
      return d;
    });
  }

  function getWeekKey() {
    return getWeekDates()[0].toISOString().slice(0, 10);
  }

  function getTodayDayIndex() {
    var dates = getWeekDates();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    for (var i = 0; i < dates.length; i++) {
      var d = new Date(dates[i]);
      d.setHours(0, 0, 0, 0);
      if (d.getTime() === today.getTime()) return i;
    }
    return -1;
  }

  function updateWeek() {
    var dates = getWeekDates();
    var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var label = dates[0].getDate() + ' ' + m[dates[0].getMonth()] + ' – ' + dates[6].getDate() + ' ' + m[dates[6].getMonth()] + ' ' + dates[6].getFullYear();
    var month = MONTHS[dates[0].getMonth()];

    $('week-label').textContent = label;
    $('month-label').textContent = month;
    if ($('week-label-t')) $('week-label-t').textContent = label;
    if ($('month-label-t')) $('month-label-t').textContent = month;

    dates.forEach(function (d, i) {
      var hdr = FULL_DAYS[i] + ' ' + d.getDate();
      var hdEl = $('hd-' + i);
      var htEl = $('ht-' + i);
      if (hdEl) hdEl.textContent = hdr;
      if (htEl) htEl.textContent = hdr;
    });

    loadWeekData();
  }

  function changeWeek(dir) {
    state.weekOffset += dir;
    updateWeek();
    if (state.monthlyView) loadMonthlyData();
  }

  // ── LOCAL STORAGE CACHE (instant load on revisit) ────────────
  var LS_KEY = 'wfr_cache';
  var LS_TTL = 5 * 60 * 1000; // 5 minute TTL

  function saveToLocalStorage(weekKey, data) {
    try {
      var payload = {
        weekKey: weekKey,
        staff: data.staff || state.staff,
        leave: data.leave || state.leaveRequests,
        rota: data.rota || state.currentRota,
        published: data.published != null ? data.published : state.isPublished,
        theatre: data.theatre || state.currentTheatre,
        ts: Date.now()
      };
      localStorage.setItem(LS_KEY, JSON.stringify(payload));
    } catch (e) { /* quota exceeded — ignore */ }
  }

  function loadFromLocalStorage(weekKey) {
    try {
      var raw = localStorage.getItem(LS_KEY);
      if (!raw) return null;
      var d = JSON.parse(raw);
      if (d.weekKey !== weekKey) return null;
      if (Date.now() - d.ts > LS_TTL) return null;
      return d;
    } catch (e) { return null; }
  }

  // ── DATA LOADING (Single request + cached) ──────────────────
  function loadData() {
    var wk = getWeekKey();

    // 1. Try localStorage for INSTANT render while fresh data loads
    var cached = loadFromLocalStorage(wk);
    if (cached) {
      state.staff = cached.staff || [];
      state.leaveRequests = cached.leave || [];
      state.currentRota = cached.rota || {};
      state.isPublished = cached.published || false;
      state.currentTheatre = cached.theatre || {};
      state.cache[wk] = { shifts: state.currentRota, published: state.isPublished, theatre: state.currentTheatre };
      populateLeaveStaffDropdown();
      updatePublishStatus();
      renderRota();
      renderTheatre();
      renderLeave();
      hideLoading();
    }

    // 2. Always fetch fresh data via single getInit call
    showToast(cached ? 'Refreshing...' : 'Loading data...', 'info');

    return api('getInit', { weekKey: wk }).then(function (result) {
      if (!result || result.error) {
        // Fallback: try individual calls if getInit not deployed yet
        return loadDataLegacy();
      }

      state.staff = result.staff || [];
      state.leaveRequests = result.leave || [];
      state.leaveBalances = result.leaveBalances || [];
      state.currentRota = result.rota || {};
      state.isPublished = result.published || false;
      state.currentTheatre = result.theatre || {};
      if (result.config) {
        state.settings = result.config;
        state.settingsLoaded = true;
      }
      if (result.has_pin) state.hasPinConfigured = true;

      // Cache in memory + localStorage
      state.cache[wk] = { shifts: state.currentRota, published: state.isPublished, theatre: state.currentTheatre };
      saveToLocalStorage(wk, result);

      populateLeaveStaffDropdown();
      updatePublishStatus();
      renderRota();
      renderTheatre();
      renderLeave();
      hideLoading();
      showToast('Data loaded', 'success');
    });
  }

  // Fallback if getInit isn't deployed yet (backward compatible)
  function loadDataLegacy() {
    return Promise.all([
      api('getStaff'),
      api('getLeave'),
      api('getRota', { weekKey: getWeekKey() }),
      api('getTheatre', { weekKey: getWeekKey() })
    ]).then(function (results) {
      var wk = getWeekKey();
      if (results[0] && results[0].staff) state.staff = results[0].staff;
      if (results[1] && results[1].requests) state.leaveRequests = results[1].requests;
      if (results[2]) {
        state.currentRota = results[2].shifts || {};
        state.isPublished = results[2].published || false;
      }
      if (results[3]) state.currentTheatre = results[3].schedule || {};

      state.cache[wk] = { shifts: state.currentRota, published: state.isPublished, theatre: state.currentTheatre };
      saveToLocalStorage(wk, {});

      populateLeaveStaffDropdown();
      updatePublishStatus();
      renderRota();
      renderTheatre();
      renderLeave();
      hideLoading();
      showToast('Data loaded', 'success');
    });
  }

  function loadWeekData() {
    var wk = getWeekKey();

    // Check memory cache
    if (state.cache[wk]) {
      var cached = state.cache[wk];
      state.currentRota = cached.shifts || {};
      state.isPublished = cached.published || false;
      state.currentTheatre = cached.theatre || {};
      updatePublishStatus();
      renderRota();
      renderTheatre();
      return Promise.resolve();
    }

    // Load both in parallel
    return Promise.all([
      api('getRota', { weekKey: wk }),
      api('getTheatre', { weekKey: wk })
    ]).then(function (results) {
      var rotaResult = results[0];
      var theatreResult = results[1];

      if (rotaResult) {
        state.currentRota = rotaResult.shifts || {};
        state.isPublished = rotaResult.published || false;
        updatePublishStatus();
      }
      if (theatreResult) {
        state.currentTheatre = theatreResult.schedule || {};
      }

      // Cache it
      state.cache[wk] = {
        shifts: state.currentRota,
        published: state.isPublished,
        theatre: state.currentTheatre,
      };

      renderRota();
      renderTheatre();
    });
  }

  function loadLeave() {
    return api('getLeave').then(function (result) {
      if (result && result.requests) {
        state.leaveRequests = result.requests;
        renderLeave();
      }
    });
  }

  // ── SHIFT PARSING ──────────────────────────────────────────────
  function getShiftClass(val) {
    if (!val) return 'empty';
    var v = val.toUpperCase().trim();
    if (v === 'DO' || v === 'DO R') return 'do';
    if (v.indexOf('A/L') === 0) return 'al';
    if (v.indexOf('S/L') === 0 || v === 'S/L') return 'sl';
    if (v.indexOf('C/L') === 0 || v === 'C/L') return 'cl';
    if (v === 'SICK') return 'sick';
    if (v === 'AV') return 'av';
    if (v.indexOf('WFH') >= 0) return 'wfh';
    if (v.slice(-2) === ' S' || v === 'S') return 'scrub';
    if (v.slice(-2) === ' A' || v === 'A') return 'assist';
    if (v.slice(-2) === ' W' || v === 'W') return 'ward';
    if (v.slice(-2) === ' R' || v === 'R') return 'rec';
    if (v.slice(-3) === ' CR' || v === 'CR') return 'cr';
    if (v.slice(-4) === ' OPD' || v === 'OPD') return 'opd';
    if (/\d/.test(v)) return 'time';
    return 'other';
  }

  function parseHours(str) {
    if (!str) return 0;
    var v = str.toUpperCase().trim();
    if (v === 'DO' || v === 'A/L' || v === 'AV' || v === 'OFF' || v.indexOf('DO ') === 0 || v === 'S/L' || v === 'C/L' || v === 'SICK') return 0;
    var wfh = v.match(/(\d+\.?\d*)\s*h/i);
    if (wfh) return parseFloat(wfh[1]);
    var m = str.match(/(\d{1,2}):?(\d{2})?\s*[-\u2013]\s*(\d{1,2}):?(\d{2})?/);
    if (!m) return 0;
    var s = parseInt(m[1]) + (parseInt(m[2] || '0') / 60);
    var e = parseInt(m[3]) + (parseInt(m[4] || '0') / 60);
    var diff = e - s;
    return Math.max(0, diff > 6 ? diff - 0.5 : diff);
  }

  // ── RENDERING: Summary ─────────────────────────────────────────
  function renderSummary() {
    var bar = $('summary-bar');
    if (!bar) return;
    var todayIdx = getTodayDayIndex();
    var todayDay = todayIdx >= 0 ? DAYS[todayIdx] : null;

    var working = 0, off = 0, onLeave = 0, bankIn = 0;
    state.staff.forEach(function (s) {
      var shifts = state.currentRota[s.id] || {};
      var val = todayDay ? (shifts[todayDay] || '') : '';
      var v = val.toUpperCase().trim();
      if (!v || v === '\u2014') { off++; }
      else if (v === 'DO' || v.indexOf('DO ') === 0) { off++; }
      else if (v.indexOf('A/L') === 0 || v === 'SICK' || v.indexOf('S/L') === 0 || v.indexOf('C/L') === 0) { onLeave++; }
      else {
        working++;
        if (s.group === 'bank') bankIn++;
      }
    });

    if (todayDay) {
      bar.innerHTML =
        '<div class="summary-chip working"><span class="num">' + working + '</span> Working today</div>' +
        '<div class="summary-chip bank"><span class="num">' + bankIn + '</span> Bank in</div>' +
        '<div class="summary-chip leave"><span class="num">' + onLeave + '</span> On leave</div>' +
        '<div class="summary-chip off"><span class="num">' + off + '</span> Off / Unassigned</div>';
    } else {
      bar.innerHTML = '';
    }
  }

  // ── RENDERING: Rota Grid ──────────────────────────────────────
  function renderRota() {
    var filter = $('group-filter') ? $('group-filter').value : 'all';
    var tbody = $('rota-body');
    var todayIdx = getTodayDayIndex();
    var html = '';

    var groups = {
      admin: state.staff.filter(function (s) { return s.group === 'admin'; }),
      clinical: state.staff.filter(function (s) { return s.group === 'clinical'; }),
      bank: state.staff.filter(function (s) { return s.group === 'bank'; }),
    };

    // Today class on header columns
    DAYS.forEach(function (d, i) {
      var th = $('hd-' + i);
      if (th) th.className = i === todayIdx ? 'today-col' : '';
    });

    function renderGroup(groupName, label, colorClass, staffList) {
      if (filter !== 'all' && filter !== groupName) return '';
      var h = '<tr class="section-header ' + colorClass + '"><td colspan="9">' + escapeHtml(label) + ' (' + staffList.length + ')</td></tr>';

      staffList.forEach(function (s) {
        var shifts = state.currentRota[s.id] || {};
        var totalHrs = 0;
        h += '<tr>';
        h += '<td><div class="staff-name">' + escapeHtml(s.name) + '</div><div class="staff-meta">' +
          escapeHtml(s.hours || (s.contractHours ? s.contractHours + 'h/wk' : s.department || '')) + '</div></td>';

        DAYS.forEach(function (d, di) {
          var val = shifts[d] || '';
          totalHrs += parseHours(val);
          var cls = getShiftClass(val);
          var todayCls = di === todayIdx ? ' today-col' : '';
          h += '<td class="' + todayCls + '">' +
            '<div class="shift-cell ' + cls + '" data-staff-id="' + escapeHtml(s.id) + '" data-day="' + d + '" title="' + (escapeHtml(val) || 'Click to assign') + '">' +
            (escapeHtml(val) || '\u2014') + '</div></td>';
        });

        var hrs = Math.round(totalHrs * 10) / 10;
        var contract = s.contractHours ? parseFloat(s.contractHours) : 0;
        var hrsClass = '';
        if (contract && hrs > contract + 5) hrsClass = 'hours-over';
        else if (contract && hrs > contract) hrsClass = 'hours-warn';
        h += '<td class="hours-cell"><div class="hours-val ' + hrsClass + '">' + (hrs > 0 ? hrs + 'h' : '\u2014') + '</div>' +
          (s.contractHours ? '<div class="hours-contract">/' + escapeHtml(String(s.contractHours)) + '</div>' : '') + '</td>';
        h += '</tr>';
      });

      // OPEN/CLOSE and CLEAN rows for clinical
      if (groupName === 'clinical' && (filter === 'all' || filter === 'clinical')) {
        ['OPEN/CLOSE', 'CLEAN'].forEach(function (lbl) {
          h += '<tr class="section-header ops"><td colspan="9" style="padding:6px 12px !important;font-size:12px">' + escapeHtml(lbl) + '</td></tr>';
          var id = 'ops-' + lbl.toLowerCase().replace(/\//g, '-');
          var shifts2 = state.currentRota[id] || {};
          h += '<tr>';
          h += '<td><div class="staff-name" style="color:var(--gray-500);font-size:12px">' + escapeHtml(lbl) + '</div></td>';
          DAYS.forEach(function (d, di) {
            var val2 = shifts2[d] || '';
            var todayCls2 = di === todayIdx ? ' today-col' : '';
            h += '<td class="' + todayCls2 + '"><div class="shift-cell ' + (val2 ? 'other' : 'empty') + '" data-staff-id="' + escapeHtml(id) + '" data-day="' + d + '">' + (escapeHtml(val2) || '\u2014') + '</div></td>';
          });
          h += '<td></td></tr>';
        });
      }
      return h;
    }

    // Theatre & Rooms section at top (only when showing all staff)
    if (filter === 'all') {
      html += '<tr class="section-header theatre"><td colspan="9">THEATRE &amp; ROOMS (' + ROOMS.length + ')</td></tr>';
      ROOMS.forEach(function (room) {
        var schedule = state.currentTheatre[room] || {};
        var isTheatre = room.indexOf('THEATRE') >= 0;
        html += '<tr>';
        html += '<td><div class="staff-name" style="font-size:13px;' + (isTheatre ? 'color:var(--primary)' : '') + '">' + escapeHtml(room) + '</div></td>';
        DAYS.forEach(function (d, di) {
          var val = schedule[d] || '';
          var todayCls = di === todayIdx ? ' today-col' : '';
          html += '<td class="' + todayCls + '"><div class="shift-cell ' + (val ? 'other' : 'empty') + '" data-room="' + escapeHtml(room) + '" data-day="' + d + '" title="' + (escapeHtml(val) || 'Click to assign') + '">' + (escapeHtml(val) || '\u2014') + '</div></td>';
        });
        html += '<td></td>';
        html += '</tr>';
      });
    }

    html += renderGroup('admin', 'ADMIN', 'admin', groups.admin);
    html += renderGroup('clinical', 'CLINICAL', 'clinical', groups.clinical);
    html += renderGroup('bank', 'BANK STAFF', 'bank', groups.bank);

    tbody.innerHTML = html;
    renderSummary();
  }

  // ── RENDERING: Theatre ────────────────────────────────────────
  function renderTheatre() {
    var tbody = $('theatre-body');
    if (!tbody) return;
    var html = '';
    ROOMS.forEach(function (room) {
      var schedule = state.currentTheatre[room] || {};
      var isTheatre = room.indexOf('THEATRE') >= 0;
      html += '<tr>';
      html += '<td style="font-weight:600;font-size:13px;color:var(--gray-700);' + (isTheatre ? 'background:#F0F9FF' : '') + '">' + escapeHtml(room) + '</td>';
      DAYS.forEach(function (d) {
        var val = schedule[d] || '';
        html += '<td><div class="shift-cell ' + (val ? 'other' : 'empty') + '" data-room="' + escapeHtml(room) + '" data-day="' + d + '">' + (escapeHtml(val) || '\u2014') + '</div></td>';
      });
      html += '</tr>';
    });
    tbody.innerHTML = html;
  }

  // ── RENDERING: Leave ──────────────────────────────────────────
  function countWeekdays(startStr, endStr) {
    var count = 0;
    var s = new Date(startStr + 'T00:00:00');
    var e = new Date(endStr + 'T00:00:00');
    var c = new Date(s);
    while (c <= e) {
      var dow = c.getDay();
      if (dow >= 1 && dow <= 5) count++;
      c.setDate(c.getDate() + 1);
    }
    return count;
  }

  function renderLeave() {
    var tbody = $('leave-body');
    if (!tbody) return;
    var pending = state.leaveRequests.filter(function (r) { return r.status === 'pending'; }).length;
    var badge = $('leave-badge');
    if (badge) {
      if (pending > 0) { badge.style.display = 'inline'; badge.textContent = pending; }
      else { badge.style.display = 'none'; }
    }

    // Render balance cards at top
    renderLeaveBalances();

    // Filter by status
    var filterEl = $('leave-status-filter');
    var statusFilter = filterEl ? filterEl.value : 'all';
    var filtered = state.leaveRequests.filter(function (r) {
      return statusFilter === 'all' || r.status === statusFilter;
    });

    var html = '';
    filtered.forEach(function (r) {
      var days = countWeekdays(String(r.startDate).slice(0, 10), String(r.endDate).slice(0, 10));
      html += '<tr>' +
        '<td style="font-weight:500">' + escapeHtml(r.staffName) + '</td>' +
        '<td>' + escapeHtml(r.type) + '</td>' +
        '<td>' + escapeHtml(String(r.startDate).slice(0, 10)) + '</td>' +
        '<td>' + escapeHtml(String(r.endDate).slice(0, 10)) + '</td>' +
        '<td style="text-align:center;font-weight:600">' + days + '</td>' +
        '<td style="color:var(--gray-500)">' + (escapeHtml(r.reason) || '\u2014') + '</td>' +
        '<td><span class="leave-status ' + escapeHtml(r.status) + '">' + escapeHtml(r.status) + '</span></td>' +
        '<td>' + (r.status === 'pending'
          ? '<button class="leave-btn approve" data-leave-action="approve" data-leave-id="' + escapeHtml(r.id) + '">Approve</button> ' +
            '<button class="leave-btn reject" data-leave-action="reject" data-leave-id="' + escapeHtml(r.id) + '">Reject</button>'
          : '') + '</td>' +
        '</tr>';
    });
    if (filtered.length === 0) {
      html = '<tr><td colspan="8" style="text-align:center;padding:40px;color:var(--gray-400)">No leave requests' + (statusFilter !== 'all' ? ' with status "' + statusFilter + '"' : '') + '</td></tr>';
    }
    tbody.innerHTML = html;
  }

  function renderLeaveBalances() {
    var container = $('leave-balances-container');
    if (!container) return;
    var balances = state.leaveBalances || [];
    if (balances.length === 0) {
      container.innerHTML = '';
      return;
    }

    var groups = [
      { key: 'admin', label: 'ADMIN', cls: 'admin' },
      { key: 'clinical', label: 'CLINICAL', cls: 'clinical' }
    ];

    var html = '<div class="leave-balance-section">';
    groups.forEach(function (g) {
      var members = balances.filter(function (b) { return b.group === g.key; });
      if (members.length === 0) return;
      html += '<div class="leave-group-header ' + g.cls + '">' + g.label + ' (' + members.length + ')</div>';
      html += '<div class="leave-balance-grid">';
      members.forEach(function (b) {
        var pct = b.entitlement > 0 ? Math.round((b.taken / b.entitlement) * 100) : 0;
        if (pct > 100) pct = 100;
        var barColor = pct < 60 ? '#10B981' : pct < 85 ? '#F59E0B' : '#EF4444';
        var yearLabel = b.yearStart ? b.yearStart.slice(5) + ' to ' + b.yearEnd.slice(5) : 'Calendar year';

        // Build other leave breakdown
        var otherParts = [];
        if (b.leaveBreakdown) {
          if (b.leaveBreakdown.study > 0) otherParts.push(b.leaveBreakdown.study + 'd Study');
          if (b.leaveBreakdown.sick > 0) otherParts.push(b.leaveBreakdown.sick + 'd Sick');
          if (b.leaveBreakdown.compassionate > 0) otherParts.push(b.leaveBreakdown.compassionate + 'd Compassionate');
        }
        var otherHtml = otherParts.length > 0 ? '<div class="card-other">' + otherParts.join(' &middot; ') + '</div>' : '';

        html += '<div class="leave-balance-card">' +
          '<div class="card-name">' + escapeHtml(b.staffName) + '</div>' +
          '<div class="card-year">' + yearLabel + '</div>' +
          '<div class="progress-bar"><div class="progress-fill" style="width:' + pct + '%;background:' + barColor + '"></div></div>' +
          '<div class="card-stats"><span><span class="stat-val">' + b.entitlement + '</span> entitled</span><span><span class="stat-val">' + b.taken + '</span> taken</span><span><span class="stat-val" style="color:' + barColor + '">' + b.remaining + '</span> remaining</span></div>' +
          otherHtml +
          '</div>';
      });
      html += '</div>';
    });
    html += '</div>';
    container.innerHTML = html;
  }

  function updatePublishStatus() {
    var el = $('publish-status');
    if (!el) return;
    el.textContent = state.isPublished ? '\u2713 Published' : '\u25CF Draft';
    el.className = 'status ' + (state.isPublished ? 'published' : 'draft');
  }

  // ── EDITING: Shift Cells ──────────────────────────────────────
  var shiftPopup = null;

  function showShiftPopup(staffId, day, cellEl, weekKey) {
    closeShiftPopup();

    var currentVal = cellEl.textContent === '\u2014' ? '' : cellEl.textContent;
    var overrideWk = weekKey || null;
    var rect = cellEl.getBoundingClientRect();

    var popup = document.createElement('div');
    popup.className = 'shift-popup';
    popup.id = 'active-shift-popup';

    // Presets
    var presetsDiv = document.createElement('div');
    presetsDiv.className = 'shift-popup-presets';
    SHIFT_PRESETS.forEach(function (preset) {
      var btn = document.createElement('button');
      btn.textContent = preset.label;
      btn.title = preset.value;
      btn.onclick = function (e) {
        e.stopPropagation();
        saveShiftFromPopup(staffId, day, preset.value, currentVal, overrideWk);
      };
      presetsDiv.appendChild(btn);
    });
    popup.appendChild(presetsDiv);

    // Custom input row
    var customRow = document.createElement('div');
    customRow.className = 'custom-row';
    var input = document.createElement('input');
    input.type = 'text';
    input.value = currentVal;
    input.placeholder = 'e.g. 07:00-15:00 S';
    input.onkeydown = function (e) {
      if (e.key === 'Enter') {
        saveShiftFromPopup(staffId, day, input.value, currentVal, overrideWk);
      }
      if (e.key === 'Escape') closeShiftPopup();
      if (e.key === 'Tab') {
        e.preventDefault();
        saveShiftFromPopup(staffId, day, input.value, currentVal, overrideWk);
        // Move to next day
        var dayIdx = DAYS.indexOf(day);
        if (dayIdx < 6) {
          var sel = '[data-staff-id="' + staffId + '"][data-day="' + DAYS[dayIdx + 1] + '"]';
          if (overrideWk) sel += '[data-week-key="' + overrideWk + '"]';
          var nextCell = document.querySelector(sel);
          if (nextCell) setTimeout(function () { nextCell.click(); }, 50);
        }
      }
    };
    customRow.appendChild(input);

    var clearBtn = document.createElement('button');
    clearBtn.className = 'clear-btn';
    clearBtn.textContent = 'Clear';
    clearBtn.onclick = function (e) {
      e.stopPropagation();
      saveShiftFromPopup(staffId, day, '', currentVal, overrideWk);
    };
    customRow.appendChild(clearBtn);
    popup.appendChild(customRow);

    document.body.appendChild(popup);
    shiftPopup = popup;

    // Position the popup
    var popupRect = popup.getBoundingClientRect();
    var top = rect.bottom + 4;
    var left = rect.left;
    if (top + popupRect.height > window.innerHeight) {
      top = rect.top - popupRect.height - 4;
    }
    if (left + popupRect.width > window.innerWidth) {
      left = window.innerWidth - popupRect.width - 8;
    }
    if (left < 8) left = 8;
    popup.style.top = top + 'px';
    popup.style.left = left + 'px';

    setTimeout(function () { input.focus(); input.select(); }, 30);
  }

  function closeShiftPopup() {
    var existing = $('active-shift-popup');
    if (existing) existing.remove();
    shiftPopup = null;
  }

  function saveShiftFromPopup(staffId, day, value, oldValue, overrideWeekKey) {
    closeShiftPopup();
    var wk = overrideWeekKey || getWeekKey();

    // Push to undo stack
    if (value !== oldValue) {
      state.undoStack.push({
        weekKey: wk,
        staffId: staffId,
        day: day,
        oldValue: oldValue,
        newValue: value,
      });
      if (state.undoStack.length > 30) state.undoStack.shift();
    }

    // Update local state + cache
    if (overrideWeekKey) {
      // Monthly view: update monthlyData and cache
      if (state.monthlyData[wk]) {
        if (!state.monthlyData[wk].shifts[staffId]) state.monthlyData[wk].shifts[staffId] = {};
        state.monthlyData[wk].shifts[staffId][day] = value;
      }
      if (state.cache[wk]) {
        if (!state.cache[wk].shifts) state.cache[wk].shifts = {};
        if (!state.cache[wk].shifts[staffId]) state.cache[wk].shifts[staffId] = {};
        state.cache[wk].shifts[staffId][day] = value;
      }
      // Also update currentRota if this is the current week
      if (wk === getWeekKey()) {
        if (!state.currentRota[staffId]) state.currentRota[staffId] = {};
        state.currentRota[staffId][day] = value;
        renderRota();
      }
      // Re-render the monthly view
      if (state.monthlyWeekKeys) renderMonthlyRota(state.monthlyWeekKeys);
    } else {
      // Weekly view: original behavior
      if (!state.currentRota[staffId]) state.currentRota[staffId] = {};
      state.currentRota[staffId][day] = value;
      invalidateCache(wk);
      renderRota();
    }

    // Save to backend
    apiWrite('saveShift', { data: { weekKey: wk, staffId: staffId, day: day, value: value } });

    if (value !== oldValue) {
      showToast('Shift saved', 'success', function () {
        undoLastShift();
      });
    }
  }

  function undoLastShift() {
    if (state.undoStack.length === 0) return;
    var last = state.undoStack.pop();
    if (!state.currentRota[last.staffId]) state.currentRota[last.staffId] = {};
    state.currentRota[last.staffId][last.day] = last.oldValue;
    invalidateCache(last.weekKey);
    renderRota();
    apiWrite('saveShift', { data: { weekKey: last.weekKey, staffId: last.staffId, day: last.day, value: last.oldValue } });
    showToast('Shift reverted', 'info');
  }

  function invalidateCache(weekKey) {
    delete state.cache[weekKey];
  }

  // ── EDITING: Theatre Cells ────────────────────────────────────
  function editTheatre(room, day, el) {
    var currentVal = el.textContent === '\u2014' ? '' : el.textContent;
    var td = el.parentElement;

    var input = document.createElement('input');
    input.className = 'shift-input';
    input.value = currentVal;
    input.placeholder = 'e.g. AQ GA';

    input.onblur = function () {
      if (!state.currentTheatre[room]) state.currentTheatre[room] = {};
      state.currentTheatre[room][day] = input.value;
      invalidateCache(getWeekKey());
      renderTheatre();
      renderRota();
      apiWrite('saveTheatre', { data: { weekKey: getWeekKey(), room: room, day: day, value: input.value } });
    };
    input.onkeydown = function (e) {
      if (e.key === 'Enter') input.blur();
      if (e.key === 'Escape') { renderTheatre(); renderRota(); }
    };

    td.innerHTML = '';
    td.appendChild(input);
    input.focus();
    input.select();
  }

  // ── ACTIONS ────────────────────────────────────────────────────
  function publishRota() {
    requirePin(function () {
      $('publish-modal').style.display = 'flex';
    });
  }

  function closePublishModal() {
    $('publish-modal').style.display = 'none';
  }

  function doPublish(mode) {
    closePublishModal();
    var labels = { 'all': 'Publishing & emailing all...', 'staff-only': 'Publishing & emailing staff...', 'none': 'Marking as published...' };
    showToast(labels[mode] || 'Publishing...', 'info');
    apiWrite('publishRota', { weekKey: getWeekKey(), emailMode: mode }).then(function (result) {
      if (result && result.success) {
        state.isPublished = true;
        invalidateCache(getWeekKey());
        updatePublishStatus();
        var msg = 'Rota published!';
        if (mode !== 'none') {
          msg += ' ' + result.emailsSent + ' staff emails';
          if (result.managerEmailsSent) msg += ', ' + result.managerEmailsSent + ' manager emails';
        }
        showToast(msg, 'success');
      } else {
        showToast('Failed to publish: ' + (result ? result.error : 'Unknown error'), 'error');
      }
    });
  }

  function copyLastWeek() {
    var prevOffset = state.weekOffset - 1;
    var today = new Date();
    var prevMonday = new Date(today);
    prevMonday.setDate(today.getDate() - ((today.getDay() + 6) % 7) + prevOffset * 7);
    var fromWeek = prevMonday.toISOString().slice(0, 10);
    var toWeek = getWeekKey();

    if (!confirm('Copy last week\'s rota to this week? This will overwrite any existing entries.')) return;
    requirePin(function () {
    showToast('Copying previous week...', 'info');
    apiWrite('copyWeek', { fromWeek: fromWeek, toWeek: toWeek }).then(function (result) {
      if (result && result.success) {
        invalidateCache(toWeek);
        loadWeekData().then(function () {
          showToast('Copied previous week\'s rota', 'success');
        });
      } else {
        showToast('No rota data found for previous week', 'warning');
      }
    });
    });
  }

  // ── LEAVE ─────────────────────────────────────────────────────
  function populateLeaveStaffDropdown() {
    var select = $('leave-staff');
    if (!select) return;
    select.innerHTML = state.staff.map(function (s) {
      return '<option value="' + escapeHtml(s.id) + '" data-name="' + escapeHtml(s.name) + '">' +
        escapeHtml(s.name) + (s.department ? ' (' + escapeHtml(s.department) + ')' : '') + '</option>';
    }).join('');
  }

  // ── STAFF SCHEDULE LOOKUP ──────────────────────────────────
  function showScheduleLookup() {
    var select = $('schedule-lookup-staff');
    var content = $('schedule-lookup-content');
    content.innerHTML = '<p style="text-align:center;color:var(--gray-400);padding:40px 0">Select a staff member above to view their upcoming schedule.</p>';

    // Populate dropdown sorted by group then name, with optgroups
    var sorted = state.staff.slice().sort(function (a, b) {
      var groupOrder = { admin: 0, clinical: 1, bank: 2 };
      var ga = groupOrder[a.group] || 9;
      var gb = groupOrder[b.group] || 9;
      if (ga !== gb) return ga - gb;
      return a.name.localeCompare(b.name);
    });

    var html = '<option value="">-- Choose a staff member --</option>';
    var lastGroup = '';
    sorted.forEach(function (s) {
      var groupLabel = (s.group || '').charAt(0).toUpperCase() + (s.group || '').slice(1);
      if (s.group !== lastGroup) {
        if (lastGroup) html += '</optgroup>';
        html += '<optgroup label="' + escapeHtml(groupLabel) + '">';
        lastGroup = s.group;
      }
      html += '<option value="' + escapeHtml(s.id) + '">' + escapeHtml(s.name) +
        (s.department ? ' (' + escapeHtml(s.department) + ')' : '') + '</option>';
    });
    if (lastGroup) html += '</optgroup>';
    select.innerHTML = html;

    select.onchange = function () {
      if (!select.value) {
        content.innerHTML = '<p style="text-align:center;color:var(--gray-400);padding:40px 0">Select a staff member above to view their upcoming schedule.</p>';
        return;
      }
      loadScheduleLookup(select.value);
    };

    $('schedule-lookup-modal').style.display = 'flex';
  }

  function closeScheduleLookup() {
    $('schedule-lookup-modal').style.display = 'none';
  }

  function loadScheduleLookup(staffId) {
    var content = $('schedule-lookup-content');
    var staffMember = state.staff.find(function (s) { return s.id === staffId; });
    if (!staffMember) return;

    content.innerHTML = '<div style="text-align:center;padding:40px 0"><div class="spinner"></div><p style="margin-top:8px;color:var(--gray-400);font-size:13px">Loading schedule...</p></div>';

    // Calculate 5 week keys starting from current week (always from today, ignoring weekOffset)
    var weekKeys = [];
    for (var w = 0; w < 5; w++) {
      var today = new Date();
      var monday = new Date(today);
      monday.setDate(today.getDate() - ((today.getDay() + 6) % 7) + w * 7);
      weekKeys.push(monday.toISOString().slice(0, 10));
    }

    // Fetch any weeks not already in cache
    var fetches = [];
    weekKeys.forEach(function (wk) {
      if (!state.cache[wk]) {
        fetches.push(
          api('getRota', { weekKey: wk }).then(function (r) {
            if (r) {
              state.cache[wk] = { shifts: r.shifts || {}, published: r.published || false, theatre: {} };
            }
          })
        );
      }
    });

    Promise.all(fetches).then(function () {
      renderScheduleLookup(staffMember, weekKeys);
    });
  }

  function renderScheduleLookup(staffMember, weekKeys) {
    var content = $('schedule-lookup-content');
    var weekM = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    // Staff info bar
    var html = '<div class="schedule-lookup-info">';
    if (staffMember.group) html += '<span>' + escapeHtml(staffMember.group.charAt(0).toUpperCase() + staffMember.group.slice(1)) + '</span>';
    if (staffMember.hours) html += '<span>' + escapeHtml(staffMember.hours) + '</span>';
    if (staffMember.contractHours) html += '<span>' + staffMember.contractHours + 'h/wk</span>';
    if (staffMember.department) html += '<span>' + escapeHtml(staffMember.department) + '</span>';
    html += '</div>';

    weekKeys.forEach(function (wk, w) {
      var monday = new Date(wk + 'T00:00:00');
      var sunday = new Date(monday);
      sunday.setDate(monday.getDate() + 6);
      var label = monday.getDate() + ' ' + weekM[monday.getMonth()] + ' \u2013 ' + sunday.getDate() + ' ' + weekM[sunday.getMonth()];
      var isCurrent = w === 0;

      var cached = state.cache[wk];
      var shifts = cached ? (cached.shifts[staffMember.id] || {}) : {};

      var totalHrs = 0;
      DAYS.forEach(function (d) { totalHrs += parseHours(shifts[d] || ''); });
      totalHrs = Math.round(totalHrs * 10) / 10;

      html += '<div class="staff-view-week">';
      html += '<h3>' + (isCurrent ? '\uD83D\uDCCD ' : '') + escapeHtml(label) + (isCurrent ? ' (This Week)' : '') +
        (totalHrs > 0 ? ' <span style="float:right;font-weight:500;font-size:12px;color:var(--gray-500)">' + totalHrs + 'h</span>' : '') + '</h3>';
      html += '<div class="days">';

      DAYS.forEach(function (d, i) {
        var dt = new Date(monday);
        dt.setDate(monday.getDate() + i);
        var val = shifts[d] || '';
        var cls = getShiftClass(val);
        html += '<div class="day-cell">' +
          '<div class="day-label">' + FULL_DAYS[i].slice(0, 3) + ' ' + dt.getDate() + '</div>' +
          '<div class="shift-cell ' + cls + '" style="min-height:32px;cursor:default">' + (escapeHtml(val) || '\u2014') + '</div></div>';
      });

      html += '</div></div>';
    });

    content.innerHTML = html;
  }

  function showLeaveModal() { $('leave-modal').style.display = 'flex'; }
  function closeLeaveModal() { $('leave-modal').style.display = 'none'; }

  function submitLeave() {
    var select = $('leave-staff');
    var staffId = select.value;
    var staffName = select.options[select.selectedIndex].dataset.name;
    var data = {
      staffId: staffId,
      staffName: staffName,
      type: $('leave-type').value,
      startDate: $('leave-start').value,
      endDate: $('leave-end').value,
      reason: $('leave-reason').value,
    };
    if (!data.startDate || !data.endDate) { showToast('Please select dates', 'warning'); return; }

    closeLeaveModal();
    showToast('Submitting leave request...', 'info');
    apiWrite('submitLeave', { data: data }).then(function (result) {
      if (result && result.success) {
        showToast('Leave request submitted & managers notified', 'success');
        loadLeave();
      } else {
        showToast('Failed to submit: ' + (result ? result.error : 'Unknown'), 'error');
      }
    });
  }

  function approveLeave(id) {
    requirePin(function () {
      showToast('Approving...', 'info');
      apiWrite('approveLeave', { id: id }).then(function (result) {
        if (result && result.success) {
          showToast('Leave approved \u2014 shifts added to rota', 'success');
          // Invalidate rota cache so auto-inserted shifts appear
          state.cache = {};
          renderRota();
        } else {
          showToast('Failed: ' + (result ? result.error : 'Unknown'), 'error');
        }
        // Refresh leave requests and balances
        loadLeave();
        refreshLeaveBalances();
      });
    });
  }

  function rejectLeave(id) {
    if (!confirm('Reject this leave request?')) return;
    requirePin(function () {
      apiWrite('rejectLeave', { id: id }).then(function (result) {
        if (result && result.success) {
          showToast('Leave rejected \u2014 staff member notified', 'warning');
        } else {
          showToast('Failed: ' + (result ? result.error : 'Unknown'), 'error');
        }
        loadLeave();
        refreshLeaveBalances();
      });
    });
  }

  function refreshLeaveBalances() {
    api('getLeaveBalances').then(function (result) {
      if (result && result.balances) {
        state.leaveBalances = result.balances;
        renderLeaveBalances();
      }
    });
  }

  // ── EMAIL LOG ─────────────────────────────────────────────────
  function loadEmailLog() {
    api('getEmailLog').then(function (result) {
      var container = $('email-log-container');
      if (!result || !result.logs || result.logs.length === 0) {
        container.innerHTML = '<div style="background:#fff;border-radius:var(--radius-lg);border:1px solid var(--gray-200);padding:50px;text-align:center"><p style="color:var(--gray-500)">No emails sent yet. Publish a rota to notify staff.</p></div>';
        return;
      }
      var html = '<table class="data-table"><thead><tr><th>Time</th><th>Type</th><th>Recipient</th><th>Details</th></tr></thead><tbody>';
      result.logs.slice(0, 50).forEach(function (l) {
        html += '<tr><td style="color:var(--gray-500);font-size:12px">' + escapeHtml(new Date(l.timestamp).toLocaleString()) +
          '</td><td style="font-weight:500">' + escapeHtml(l.type) +
          '</td><td>' + escapeHtml(l.recipient) +
          '</td><td style="color:var(--gray-500)">' + escapeHtml(l.details) + '</td></tr>';
      });
      html += '</tbody></table>';
      container.innerHTML = html;
    });
  }

  // ── STAFF MANAGEMENT ──────────────────────────────────────────
  function showStaffModal(id) {
    state.editingStaffId = null;
    $('staff-modal-title').textContent = 'Add Staff Member';
    $('staff-name').value = '';
    $('staff-group').value = '';
    $('staff-department').value = '';
    $('staff-hours-type').value = '';
    $('staff-contract-hours').value = '';
    $('staff-email').value = '';
    $('staff-phone').value = '';
    $('staff-start-date').value = '';
    $('staff-leave-entitlement').value = '';
    $('staff-modal').style.display = 'flex';
  }

  function closeStaffModal() {
    $('staff-modal').style.display = 'none';
    state.editingStaffId = null;
  }

  function editStaffMember(id) {
    var s = state.staff.find(function (x) { return x.id === id; });
    if (!s) return;
    state.editingStaffId = id;
    $('staff-modal-title').textContent = 'Edit Staff Member';
    $('staff-name').value = s.name || '';
    $('staff-group').value = s.group || '';
    $('staff-department').value = s.department || '';
    $('staff-hours-type').value = s.hours || '';
    $('staff-contract-hours').value = s.contractHours || '';
    $('staff-email').value = s.email || '';
    $('staff-phone').value = s.phone || '';
    $('staff-start-date').value = s.startdate || '';
    $('staff-leave-entitlement').value = s.leaveentitlement || '';
    $('staff-modal').style.display = 'flex';
  }

  function submitStaff() {
    var data = {
      name: $('staff-name').value.trim(),
      group: $('staff-group').value,
      department: $('staff-department').value.trim(),
      hours: $('staff-hours-type').value,
      contractHours: parseFloat($('staff-contract-hours').value) || 0,
      email: $('staff-email').value.trim(),
      phone: $('staff-phone').value.trim(),
      startdate: $('staff-start-date').value,
      leaveentitlement: parseInt($('staff-leave-entitlement').value) || 0,
    };
    if (!data.name || !data.group) { showToast('Name and Group are required', 'warning'); return; }

    var editingId = state.editingStaffId; // Save before closeStaffModal clears it
    closeStaffModal();
    requirePin(function () {
      if (editingId) {
        data.id = editingId;
        showToast('Updating staff...', 'info');
        apiWrite('updateStaff', { data: data }).then(function (result) {
          if (result && result.success) {
            return api('getStaff').then(function (r2) {
              if (r2 && r2.staff) { state.staff = r2.staff; populateLeaveStaffDropdown(); }
              renderStaffTable();
              showToast('Staff updated', 'success');
            });
          } else {
            showToast('Failed to update: ' + (result ? result.error : 'Unknown'), 'error');
          }
        });
      } else {
        showToast('Adding staff...', 'info');
        apiWrite('addStaff', { data: data }).then(function (result) {
          if (result && result.success) {
            return api('getStaff').then(function (r2) {
              if (r2 && r2.staff) { state.staff = r2.staff; populateLeaveStaffDropdown(); }
              renderStaffTable();
              showToast('Staff member added', 'success');
            });
          } else {
            showToast('Failed to add: ' + (result ? result.error : 'Unknown'), 'error');
          }
        });
      }
    });
  }

  function deleteStaffMember(id) {
    var s = state.staff.find(function (x) { return x.id === id; });
    if (!confirm('Delete ' + (s ? s.name : id) + '? This cannot be undone.')) return;
    requirePin(function () {
      showToast('Deleting...', 'info');
      apiWrite('deleteStaff', { id: id }).then(function (result) {
        if (result && result.success) {
          return api('getStaff').then(function (r2) {
            if (r2 && r2.staff) { state.staff = r2.staff; populateLeaveStaffDropdown(); }
            renderStaffTable();
            showToast('Staff member deleted', 'success');
          });
        } else {
          showToast('Failed to delete', 'error');
        }
      });
    });
  }

  function sendMonthlyEmail(staffId) {
    var s = state.staff.find(function (x) { return x.id === staffId; });
    if (!s) return;
    if (!s.email) { showToast(s.name + ' has no email address', 'warning'); return; }
    if (!confirm('Send ' + s.name + ' their rota for the next 5 weeks?')) return;
    showToast('Sending monthly rota to ' + s.name + '...', 'info');
    apiWrite('publishMonthlyRota', { staffId: staffId, startWeekKey: getWeekKey() }).then(function (result) {
      if (result && result.success) {
        showToast('Monthly rota emailed to ' + s.name, 'success');
      } else {
        showToast('Failed: ' + (result ? result.error : 'Unknown'), 'error');
      }
    });
  }

  function renderStaffTable() {
    var filter = $('staff-group-filter') ? $('staff-group-filter').value : 'all';
    var search = state.staffSearch.toLowerCase();
    var tbody = $('staff-body');
    if (!tbody) return;
    var html = '';

    var groupsList = [
      { key: 'admin', label: 'ADMIN' },
      { key: 'clinical', label: 'CLINICAL' },
      { key: 'bank', label: 'BANK' },
    ];

    groupsList.forEach(function (g) {
      if (filter !== 'all' && filter !== g.key) return;
      var list = state.staff.filter(function (s) {
        if (s.group !== g.key) return false;
        if (search && s.name.toLowerCase().indexOf(search) < 0 && (s.email || '').toLowerCase().indexOf(search) < 0 && (s.department || '').toLowerCase().indexOf(search) < 0) return false;
        return true;
      });
      if (list.length === 0) return;
      html += '<tr><td colspan="10" style="padding:10px 14px;font-weight:700;font-size:12px;color:var(--primary);background:var(--gray-50);letter-spacing:0.5px;border-bottom:2px solid var(--primary)">' + escapeHtml(g.label) + ' (' + list.length + ')</td></tr>';
      list.forEach(function (s) {
        html += '<tr>' +
          '<td style="font-weight:500">' + escapeHtml(s.name) + '</td>' +
          '<td><span style="background:var(--gray-100);padding:3px 8px;border-radius:4px;font-size:11px;text-transform:capitalize">' + escapeHtml(s.group) + '</span></td>' +
          '<td style="color:var(--gray-500);font-size:13px">' + (escapeHtml(s.department) || '\u2014') + '</td>' +
          '<td style="font-size:13px">' + (escapeHtml(s.hours) || '\u2014') + '</td>' +
          '<td style="text-align:center;font-size:13px;font-weight:600">' + (s.contractHours ? escapeHtml(String(s.contractHours)) : '\u2014') + '</td>' +
          '<td style="font-size:13px;color:var(--gray-500)">' + (s.startdate ? escapeHtml(s.startdate) : '\u2014') + '</td>' +
          '<td style="text-align:center;font-size:13px;font-weight:600">' + (s.leaveentitlement ? escapeHtml(String(s.leaveentitlement)) : '\u2014') + '</td>' +
          '<td style="font-size:13px">' + (s.email ? '<a href="mailto:' + escapeHtml(s.email) + '" style="color:var(--primary);text-decoration:none">' + escapeHtml(s.email) + '</a>' : '<span style="color:var(--gray-300)">Not set</span>') + '</td>' +
          '<td style="font-size:13px">' + (escapeHtml(s.phone) || '<span style="color:var(--gray-300)">Not set</span>') + '</td>' +
          '<td style="white-space:nowrap">' +
          (s.email ? '<button class="staff-action-btn email-month" data-staff-email-month="' + escapeHtml(s.id) + '" title="Email 5-week rota">Email Month</button>' : '') +
          '<button class="staff-action-btn edit" data-staff-edit="' + escapeHtml(s.id) + '">Edit</button>' +
          '<button class="staff-action-btn delete" data-staff-delete="' + escapeHtml(s.id) + '">Delete</button>' +
          '</td></tr>';
      });
    });

    if (html === '') html = '<tr><td colspan="10" style="text-align:center;padding:40px;color:var(--gray-400)">No staff members found</td></tr>';
    tbody.innerHTML = html;
  }

  // ── SETTINGS ──────────────────────────────────────────────────
  function loadSettings() {
    if (state.settingsLoaded) { renderManagerList(); renderLegend(); return; }
    api('getConfig').then(function (result) {
      if (result && result.config) state.settings = result.config;
      state.settingsLoaded = true;
      renderManagerList();
      renderLegend();
      var cb = $('auto-email-pref');
      if (cb) cb.checked = state.settings.auto_email !== false;
    });
  }

  function renderManagerList() {
    var container = $('manager-list');
    if (!container) return;
    var emails = state.settings.manager_emails || [];
    if (emails.length === 0) {
      container.innerHTML = '<p style="font-size:13px;color:var(--gray-400);padding:8px 0">No managers added yet</p>';
      return;
    }
    container.innerHTML = emails.map(function (email) {
      return '<div class="manager-row"><span>' + escapeHtml(email) + '</span><button data-remove-manager="' + escapeHtml(email) + '" title="Remove">\u2715</button></div>';
    }).join('');
  }

  function addManager() {
    var input = $('manager-email-input');
    var email = input.value.trim().toLowerCase();
    if (!email) { showToast('Enter an email address', 'warning'); return; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { showToast('Invalid email address', 'warning'); return; }
    if ((state.settings.manager_emails || []).indexOf(email) >= 0) { showToast('Manager already added', 'warning'); return; }
    if (!state.settings.manager_emails) state.settings.manager_emails = [];
    state.settings.manager_emails.push(email);
    input.value = '';
    saveSettings();
    renderManagerList();
    showToast('Manager added', 'success');
  }

  function removeManager(email) {
    state.settings.manager_emails = (state.settings.manager_emails || []).filter(function (e) { return e !== email; });
    saveSettings();
    renderManagerList();
    showToast('Manager removed', 'success');
  }

  function saveSettings() {
    apiWrite('saveConfig', { config: state.settings });
  }

  function saveNotificationPrefs() {
    state.settings.auto_email = $('auto-email-pref').checked;
    saveSettings();
    showToast('Preference saved', 'success');
  }

  // ── MONTHLY VIEW ─────────────────────────────────────────────
  function toggleMonthlyView() {
    state.monthlyView = !state.monthlyView;
    var btn = $('monthly-toggle');
    var weeklyGrid = $('weekly-rota-section');
    var monthlyGrid = $('monthly-rota-section');
    var weekNav = $('week-nav-controls');

    if (state.monthlyView) {
      btn.innerHTML = '\u2190 Back to Weekly';
      weeklyGrid.style.display = 'none';
      monthlyGrid.style.display = 'block';
      if (weekNav) weekNav.style.display = 'none';
      loadMonthlyData();
    } else {
      btn.innerHTML = '&#x1F4C6; Monthly View';
      weeklyGrid.style.display = 'block';
      monthlyGrid.style.display = 'none';
      if (weekNav) weekNav.style.display = '';
    }
  }

  function getMonthlyWeekKeys() {
    var weekKeys = [];
    for (var w = 0; w < 5; w++) {
      var today = new Date();
      var monday = new Date(today);
      monday.setDate(today.getDate() - ((today.getDay() + 6) % 7) + (state.weekOffset + w) * 7);
      weekKeys.push(monday.toISOString().slice(0, 10));
    }
    return weekKeys;
  }

  function loadMonthlyData() {
    var weekKeys = getMonthlyWeekKeys();
    state.monthlyWeekKeys = weekKeys;

    // Populate from cache what we already have
    weekKeys.forEach(function (wk) {
      if (state.cache[wk]) state.monthlyData[wk] = state.cache[wk];
    });

    // Render immediately with whatever we have (current week is always cached)
    renderMonthlyRota(weekKeys);

    // Fetch any missing weeks
    var fetches = [];
    weekKeys.forEach(function (wk) {
      if (!state.monthlyData[wk]) {
        fetches.push(
          api('getRota', { weekKey: wk }).then(function (r) {
            if (r) {
              state.monthlyData[wk] = { shifts: r.shifts || {}, published: r.published || false };
              state.cache[wk] = state.monthlyData[wk];
            }
          })
        );
      }
    });

    if (fetches.length > 0) {
      showToast('Loading remaining weeks...', 'info');
      Promise.all(fetches).then(function () {
        renderMonthlyRota(weekKeys);
        showToast('Monthly view loaded', 'success');
      });
    }
  }

  function renderMonthlyRota(weekKeys) {
    var container = $('monthly-rota-container');
    if (!container) return;
    var filter = $('group-filter') ? $('group-filter').value : 'all';
    var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    // Pre-compute groups (same staff for all weeks)
    var groups = {
      admin: state.staff.filter(function (s) { return s.group === 'admin'; }),
      clinical: state.staff.filter(function (s) { return s.group === 'clinical'; }),
      bank: state.staff.filter(function (s) { return s.group === 'bank'; }),
    };

    // Compute dates for each week (for dynamic header update)
    var weekDates = weekKeys.map(function (wk) {
      var monday = new Date(wk + 'T00:00:00');
      return DAYS.map(function (d, i) {
        var dt = new Date(monday);
        dt.setDate(monday.getDate() + i);
        return dt.getDate();
      });
    });

    // Build single combined table
    var html = '<div class="rota-grid monthly-combined">';
    html += '<div class="rota-grid-scroll">';
    html += '<table><thead id="monthly-thead"><tr><th>Staff</th>';

    // Initial header shows first week's dates
    var firstDates = weekDates[0] || [];
    DAYS.forEach(function (d, i) {
      html += '<th>' + d + ' ' + (firstDates[i] || '') + '</th>';
    });
    html += '<th>HRS</th></tr></thead><tbody>';

    weekKeys.forEach(function (wk, wIdx) {
      var monday = new Date(wk + 'T00:00:00');
      var sunday = new Date(monday);
      sunday.setDate(monday.getDate() + 6);
      var weekLabel = monday.getDate() + ' ' + m[monday.getMonth()] + ' \u2013 ' + sunday.getDate() + ' ' + m[sunday.getMonth()];
      var isCurrent = wIdx === 0 && state.weekOffset === 0;
      var weekData = state.monthlyData[wk];
      var hasData = !!weekData;
      if (!weekData) weekData = { shifts: {}, published: false };

      // Week divider row (sticky below thead)
      html += '<tr class="monthly-week-divider' + (isCurrent ? ' current' : '') + '" data-dates="' + weekDates[wIdx].join(',') + '">';
      html += '<td colspan="9">';
      html += '<span class="monthly-week-label">' + escapeHtml(weekLabel) + (isCurrent ? ' (This Week)' : '') + '</span>';
      html += '<span class="monthly-week-status ' + (weekData.published ? 'published' : 'draft') + '">' + (weekData.published ? '\u2713 Published' : '\u25CF Draft') + '</span>';
      html += '</td></tr>';

      if (!hasData) {
        html += '<tr><td colspan="9" style="padding:30px;text-align:center;color:var(--gray-400);font-size:13px"><div class="spinner"></div><p style="margin-top:8px">Loading...</p></td></tr>';
        return;
      }

      function renderMonthlyGroup(groupName, label, colorClass, staffList) {
        if (filter !== 'all' && filter !== groupName) return '';
        var h = '<tr class="section-header ' + colorClass + '"><td colspan="9">' + escapeHtml(label) + ' (' + staffList.length + ')</td></tr>';
        staffList.forEach(function (s) {
          var shifts = weekData.shifts[s.id] || {};
          var totalHrs = 0;
          h += '<tr><td><div class="staff-name">' + escapeHtml(s.name) + '</div></td>';
          DAYS.forEach(function (d) {
            var val = shifts[d] || '';
            totalHrs += parseHours(val);
            var cls = getShiftClass(val);
            h += '<td><div class="shift-cell ' + cls + '" data-staff-id="' + escapeHtml(s.id) + '" data-day="' + d + '" data-week-key="' + escapeHtml(wk) + '" title="' + (escapeHtml(val) || 'Click to assign') + '">' + (escapeHtml(val) || '\u2014') + '</div></td>';
          });
          var hrs = Math.round(totalHrs * 10) / 10;
          h += '<td class="hours-cell"><div class="hours-val">' + (hrs > 0 ? hrs + 'h' : '\u2014') + '</div></td></tr>';
        });
        return h;
      }

      html += renderMonthlyGroup('admin', 'ADMIN', 'admin', groups.admin);
      html += renderMonthlyGroup('clinical', 'CLINICAL', 'clinical', groups.clinical);
      html += renderMonthlyGroup('bank', 'BANK STAFF', 'bank', groups.bank);
    });

    html += '</tbody></table></div></div>';
    container.innerHTML = html;

    // Set up IntersectionObserver to update thead dates as user scrolls between weeks
    var scrollEl = container.querySelector('.rota-grid-scroll');
    var theadRow = container.querySelector('#monthly-thead tr');
    var dividers = container.querySelectorAll('.monthly-week-divider');

    if (scrollEl && theadRow && dividers.length > 0) {
      var theadCells = theadRow.querySelectorAll('th');

      function updateHeaderDates() {
        var scrollRect = scrollEl.getBoundingClientRect();
        var topDivider = null;
        dividers.forEach(function (d) {
          var rect = d.getBoundingClientRect();
          if (rect.top <= scrollRect.top + 80) topDivider = d;
        });
        if (!topDivider) topDivider = dividers[0];
        if (topDivider) {
          var dates = topDivider.dataset.dates.split(',');
          DAYS.forEach(function (d, i) {
            theadCells[i + 1].textContent = d + ' ' + dates[i];
          });
        }
      }

      scrollEl.addEventListener('scroll', updateHeaderDates);
    }
  }

  function emailMonthlyOverview() {
    requirePin(function () {
      if (!confirm('Email the 5-week rota overview to all managers?')) return;
      showToast('Sending monthly overview to managers...', 'info');
      apiWrite('publishMonthlyOverview', { startWeekKey: getWeekKey() }).then(function (result) {
        if (result && result.success) {
          showToast('Monthly overview sent to ' + (result.count || 0) + ' manager(s)', 'success');
        } else {
          showToast('Failed: ' + (result ? result.error : 'Unknown'), 'error');
        }
      });
    });
  }

  // ── LEGEND ────────────────────────────────────────────────────
  function renderLegend() {
    var grids = ['settings-legend-grid', 'rota-legend-grid'];
    grids.forEach(function (gridId) {
      var grid = $(gridId);
      if (!grid) return;
      grid.innerHTML = LEGEND_ITEMS.map(function (item) {
        return '<div class="legend-item"><div class="shift-cell ' + item.cls + '" style="min-height:28px;font-size:11px">' + escapeHtml(item.label) + '</div><p>' + escapeHtml(item.desc) + '</p></div>';
      }).join('');
    });
  }

  function toggleLegendPanel() {
    state.legendVisible = !state.legendVisible;
    var panel = $('legend-panel');
    var toggle = $('legend-toggle');
    panel.style.display = state.legendVisible ? 'block' : 'none';
    toggle.textContent = state.legendVisible ? '\u25BC Hide Colour Legend' : '\u25B6 Show Colour Legend';
    if (state.legendVisible) renderLegend();
  }

  // ── CSV EXPORT ────────────────────────────────────────────────
  function exportCSV() {
    var dates = getWeekDates();
    var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var weekLabel = dates[0].getDate() + ' ' + m[dates[0].getMonth()] + ' - ' + dates[6].getDate() + ' ' + m[dates[6].getMonth()] + ' ' + dates[6].getFullYear();

    var rows = [];
    rows.push(['Staff', 'Group'].concat(DAYS).concat(['Total Hours']).join(','));

    // Theatre & Rooms section
    rows.push('');
    rows.push('"THEATRE & ROOMS",,,,,,,,');
    ROOMS.forEach(function (room) {
      var schedule = state.currentTheatre[room] || {};
      var row = ['"' + room + '"', '"theatre"'];
      DAYS.forEach(function (d) {
        row.push('"' + (schedule[d] || '') + '"');
      });
      row.push('');
      rows.push(row.join(','));
    });
    rows.push('');

    state.staff.forEach(function (s) {
      var shifts = state.currentRota[s.id] || {};
      var totalHrs = 0;
      var row = ['"' + s.name + '"', s.group];
      DAYS.forEach(function (d) {
        var val = shifts[d] || '';
        totalHrs += parseHours(val);
        row.push('"' + val + '"');
      });
      row.push(Math.round(totalHrs * 10) / 10);
      rows.push(row.join(','));
    });

    var csv = rows.join('\n');
    var blob = new Blob([csv], { type: 'text/csv' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'rota-' + getWeekKey() + '.csv';
    a.click();
    URL.revokeObjectURL(url);
    showToast('CSV exported', 'success');
  }

  // ── STAFF PERSONAL VIEW ───────────────────────────────────────
  function checkStaffView() {
    var params = new URLSearchParams(window.location.search);
    var staffName = params.get('staff');
    if (!staffName) return false;

    state.staffViewMode = true;
    state.staffViewName = staffName;
    return true;
  }

  function renderStaffView() {
    var appEl = $('app');
    var staffMember = state.staff.find(function (s) {
      return s.name.toLowerCase() === state.staffViewName.toLowerCase();
    });

    if (!staffMember) {
      appEl.innerHTML =
        '<div class="staff-view-header"><h1>Staff not found</h1><a href="?" class="back-link">\u2190 Full Rota</a></div>' +
        '<div class="content"><p style="padding:40px;text-align:center;color:var(--gray-500)">No staff member named "' + escapeHtml(state.staffViewName) + '" was found.</p></div>';
      return;
    }

    var html = '<div class="staff-view-header">' +
      '<div><h1>' + escapeHtml(staffMember.name) + '\'s Rota</h1>' +
      '<div style="color:rgba(255,255,255,0.7);font-size:13px;margin-top:2px">' +
      escapeHtml((staffMember.group || '').charAt(0).toUpperCase() + (staffMember.group || '').slice(1)) +
      (staffMember.department ? ' \u2014 ' + escapeHtml(staffMember.department) : '') +
      (staffMember.contractHours ? ' \u2014 ' + staffMember.contractHours + 'h/wk' : '') +
      '</div></div>' +
      '<a href="?" class="back-link">\u2190 Full Rota View</a></div>';

    html += '<div class="content" style="max-width:800px">';

    // Show 5 weeks: current + 4 upcoming
    for (var w = 0; w < 5; w++) {
      var today = new Date();
      var monday = new Date(today);
      monday.setDate(today.getDate() - ((today.getDay() + 6) % 7) + w * 7);
      var weekKey = monday.toISOString().slice(0, 10);
      var weekM = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var sunday = new Date(monday);
      sunday.setDate(monday.getDate() + 6);
      var label = monday.getDate() + ' ' + weekM[monday.getMonth()] + ' \u2013 ' + sunday.getDate() + ' ' + weekM[sunday.getMonth()];
      var isCurrent = w === 0;

      html += '<div class="staff-view-week">';
      html += '<h3>' + (isCurrent ? '\uD83D\uDCCD ' : '') + escapeHtml(label) + (isCurrent ? ' (This Week)' : '') + '</h3>';
      html += '<div class="days">';

      DAYS.forEach(function (d, i) {
        var dt = new Date(monday);
        dt.setDate(monday.getDate() + i);
        var cached = state.cache[weekKey];
        var shifts = cached ? (cached.shifts[staffMember.id] || {}) : {};
        var val = shifts[d] || '';
        var cls = getShiftClass(val);
        html += '<div class="day-cell">' +
          '<div class="day-label">' + FULL_DAYS[i].slice(0, 3) + ' ' + dt.getDate() + '</div>' +
          '<div class="shift-cell ' + cls + '" style="min-height:32px;cursor:default">' + (escapeHtml(val) || '\u2014') + '</div></div>';
      });

      html += '</div></div>';
    }

    html += '</div>';
    appEl.innerHTML = html;
  }

  function loadStaffViewData() {
    showToast('Loading your rota...', 'info');
    return api('getStaff').then(function (result) {
      if (result && result.staff) state.staff = result.staff;

      // Load 5 weeks of data in parallel
      var promises = [];
      for (var w = 0; w < 5; w++) {
        (function (offset) {
          var today = new Date();
          var monday = new Date(today);
          monday.setDate(today.getDate() - ((today.getDay() + 6) % 7) + offset * 7);
          var wk = monday.toISOString().slice(0, 10);
          promises.push(
            api('getRota', { weekKey: wk }).then(function (r) {
              if (r) {
                state.cache[wk] = {
                  shifts: r.shifts || {},
                  published: r.published || false,
                  theatre: {},
                };
              }
            })
          );
        })(w);
      }

      return Promise.all(promises);
    }).then(function () {
      hideLoading();
      renderStaffView();
      showToast('Rota loaded', 'success');
    });
  }

  // ── TAB SWITCHING ─────────────────────────────────────────────
  function switchTab(name) {
    if (name === 'settings') {
      requirePin(function () { doSwitchTab('settings'); });
      return;
    }
    doSwitchTab(name);
  }

  function doSwitchTab(name) {
    state.currentTab = name;
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.remove('active'); });
    var activeTab = document.querySelector('.tab[data-tab="' + name + '"]');
    if (activeTab) activeTab.classList.add('active');

    ['rota', 'theatre', 'leave', 'emails', 'staff', 'settings'].forEach(function (t) {
      var el = $('tab-' + t);
      if (el) el.style.display = t === name ? 'block' : 'none';
    });

    if (name === 'emails') loadEmailLog();
    if (name === 'staff') renderStaffTable();
    if (name === 'settings') loadSettings();
  }

  // ── EVENT DELEGATION ──────────────────────────────────────────
  function setupEventDelegation() {
    // Rota grid clicks
    $('rota-body').addEventListener('click', function (e) {
      var cell = e.target.closest('.shift-cell');
      if (!cell) return;
      var day = cell.dataset.day;
      // Theatre cell?
      var room = cell.dataset.room;
      if (room && day) {
        requirePin(function () { editTheatre(room, day, cell); });
        return;
      }
      // Staff cell
      var staffId = cell.dataset.staffId;
      if (staffId && day) {
        requirePin(function () { showShiftPopup(staffId, day, cell); });
      }
    });

    // Monthly rota grid clicks
    var monthlyContainer = $('monthly-rota-container');
    if (monthlyContainer) {
      monthlyContainer.addEventListener('click', function (e) {
        var cell = e.target.closest('.shift-cell');
        if (!cell) return;
        var staffId = cell.dataset.staffId;
        var day = cell.dataset.day;
        var weekKey = cell.dataset.weekKey;
        if (staffId && day && weekKey) {
          requirePin(function () { showShiftPopup(staffId, day, cell, weekKey); });
        }
      });
    }

    // Theatre grid clicks
    var theatreBody = $('theatre-body');
    if (theatreBody) {
      theatreBody.addEventListener('click', function (e) {
        var cell = e.target.closest('.shift-cell');
        if (!cell) return;
        var room = cell.dataset.room;
        var day = cell.dataset.day;
        if (room && day) {
          requirePin(function () { editTheatre(room, day, cell); });
        }
      });
    }

    // Leave action buttons
    var leaveBody = $('leave-body');
    if (leaveBody) {
      leaveBody.addEventListener('click', function (e) {
        var btn = e.target.closest('[data-leave-action]');
        if (!btn) return;
        var action = btn.dataset.leaveAction;
        var id = btn.dataset.leaveId;
        if (action === 'approve') approveLeave(id);
        if (action === 'reject') rejectLeave(id);
      });
    }

    // Staff table action buttons
    var staffBody = $('staff-body');
    if (staffBody) {
      staffBody.addEventListener('click', function (e) {
        var editBtn = e.target.closest('[data-staff-edit]');
        var deleteBtn = e.target.closest('[data-staff-delete]');
        var emailBtn = e.target.closest('[data-staff-email-month]');
        if (editBtn) editStaffMember(editBtn.dataset.staffEdit);
        if (deleteBtn) deleteStaffMember(deleteBtn.dataset.staffDelete);
        if (emailBtn) sendMonthlyEmail(emailBtn.dataset.staffEmailMonth);
      });
    }

    // Manager remove buttons
    var managerList = $('manager-list');
    if (managerList) {
      managerList.addEventListener('click', function (e) {
        var btn = e.target.closest('[data-remove-manager]');
        if (btn) removeManager(btn.dataset.removeManager);
      });
    }

    // Tab clicks
    document.querySelectorAll('.tab').forEach(function (tab) {
      tab.addEventListener('click', function () {
        switchTab(tab.dataset.tab);
      });
    });

    // Close shift popup on outside click
    document.addEventListener('click', function (e) {
      if (shiftPopup && !shiftPopup.contains(e.target) && !e.target.closest('.shift-cell')) {
        closeShiftPopup();
      }
    });

    // Staff search
    var searchInput = $('staff-search');
    if (searchInput) {
      searchInput.addEventListener('input', function () {
        state.staffSearch = searchInput.value;
        renderStaffTable();
      });
    }
  }

  // ── KEYBOARD NAVIGATION ───────────────────────────────────────
  function setupKeyboard() {
    document.addEventListener('keydown', function (e) {
      // Don't intercept when typing in inputs (except our shift popup)
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') return;

      if (e.key === 'ArrowLeft') { changeWeek(-1); e.preventDefault(); }
      if (e.key === 'ArrowRight') { changeWeek(1); e.preventDefault(); }
      if (e.key === 'Escape') closeShiftPopup();

      // Ctrl+Z / Cmd+Z to undo
      if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
        e.preventDefault();
        undoLastShift();
      }
    });
  }

  // ── MOBILE SWIPE ──────────────────────────────────────────────
  function setupSwipe() {
    var touchStartX = 0;
    var touchStartY = 0;
    document.addEventListener('touchstart', function (e) {
      touchStartX = e.changedTouches[0].screenX;
      touchStartY = e.changedTouches[0].screenY;
    }, { passive: true });
    document.addEventListener('touchend', function (e) {
      var dx = e.changedTouches[0].screenX - touchStartX;
      var dy = e.changedTouches[0].screenY - touchStartY;
      if (Math.abs(dx) > 100 && Math.abs(dx) > Math.abs(dy) * 2) {
        if (e.target.closest('.rota-grid-scroll')) return;
        if (e.target.closest('.shift-popup')) return;
        if (dx < 0) changeWeek(1);
        else changeWeek(-1);
      }
    }, { passive: true });
  }

  // ── DEMO MODE ─────────────────────────────────────────────────
  function loadDemoData() {
    $('config-banner').style.display = 'block';
    state.staff = [
      { id:'adm-01', name:'Janey', group:'admin', hours:'FULL TIME', contractHours:37.5 },
      { id:'adm-02', name:'Katrina', group:'admin', hours:'FULL TIME', contractHours:37.5 },
      { id:'adm-03', name:'Debbie', group:'admin', hours:'PART TIME', contractHours:22.5 },
      { id:'adm-04', name:'Danielle', group:'admin', hours:'PART TIME', contractHours:15.5 },
      { id:'adm-05', name:'Margaret', group:'admin', hours:'PART TIME', contractHours:15 },
      { id:'adm-06', name:'Lucy', group:'admin', hours:'BANK', contractHours:0 },
      { id:'cli-01', name:'Morag', group:'clinical', contractHours:37.5 },
      { id:'cli-02', name:'Islay', group:'clinical', contractHours:25 },
      { id:'cli-03', name:'Oksana', group:'clinical', contractHours:37.5 },
      { id:'cli-04', name:'Christina', group:'clinical', contractHours:37.5 },
      { id:'cli-05', name:'Cezary', group:'clinical', contractHours:37.5 },
      { id:'cli-06', name:'Piel', group:'clinical', contractHours:37.5 },
      { id:'cli-07', name:'Inez', group:'clinical', contractHours:30 },
      { id:'cli-08', name:'Laura', group:'clinical', contractHours:36 },
      { id:'cli-09', name:'Kerry Ann', group:'clinical', contractHours:33 },
      { id:'cli-10', name:'Wojtek', group:'clinical', contractHours:37.5 },
      { id:'bnk-01', name:'Gleb', group:'bank', department:'' },
      { id:'bnk-02', name:'Rona', group:'bank', department:'' },
      { id:'bnk-03', name:'Sybil', group:'bank', department:'' },
      { id:'bnk-04', name:'Gomes', group:'bank', department:'' },
      { id:'bnk-05', name:'Myron', group:'bank', department:'' },
      { id:'bnk-06', name:'Claudia', group:'bank', department:'OPD' },
      { id:'bnk-07', name:'Louise', group:'bank', department:'OPD' },
      { id:'bnk-08', name:'Sally', group:'bank', department:'OPD/Ward' },
      { id:'bnk-09', name:'Jo', group:'bank', department:'OPD' },
      { id:'bnk-10', name:'Eve', group:'bank', department:'Ward' },
      { id:'bnk-11', name:'Damien', group:'bank', department:'Ward' },
      { id:'bnk-12', name:'Draga', group:'bank', department:'Ward' },
      { id:'bnk-13', name:'Irina', group:'bank', department:'Ward' },
      { id:'bnk-14', name:'Olga', group:'bank', department:'ODP' },
      { id:'bnk-15', name:'Ashleigh', group:'bank', department:'ODP' },
      { id:'bnk-16', name:'Steve G', group:'bank', department:'ODP' },
      { id:'bnk-17', name:'Elisabeth', group:'bank', department:'TH' },
      { id:'bnk-18', name:'Clare', group:'bank', department:'Scrub/Rec' },
      { id:'bnk-19', name:'Isabella', group:'bank', department:'REC' },
      { id:'bnk-20', name:'Maxine', group:'bank', department:'ODP' },
      { id:'bnk-21', name:'Fiona B', group:'bank', department:'Scrub' },
      { id:'bnk-22', name:'Carl', group:'bank', department:'TH' },
      { id:'bnk-23', name:'Mariam', group:'bank', department:'' },
    ];
    populateLeaveStaffDropdown();
    renderRota();
    renderTheatre();
    hideLoading();
  }

  // ── INIT ──────────────────────────────────────────────────────
  function init() {
    // Check for staff personal view
    if (checkStaffView()) {
      loadStaffViewData();
      return;
    }

    updateWeek();
    setupEventDelegation();
    setupKeyboard();
    setupSwipe();

    if (API_URL !== 'YOUR_APPS_SCRIPT_URL') {
      loadData();
    } else {
      loadDemoData();
    }
  }

  // ── HELPER: Re-render active rota view ────────────────────────
  function refreshRotaView() {
    renderRota();
    if (state.monthlyView && state.monthlyWeekKeys) {
      renderMonthlyRota(state.monthlyWeekKeys);
    }
  }

  // ── EXPOSE TO WINDOW (for HTML onclick handlers) ──────────────
  window.App = {
    changeWeek: changeWeek,
    publishRota: publishRota,
    closePublishModal: closePublishModal,
    doPublish: doPublish,
    copyLastWeek: copyLastWeek,
    showLeaveModal: showLeaveModal,
    closeLeaveModal: closeLeaveModal,
    submitLeave: submitLeave,
    showStaffModal: showStaffModal,
    closeStaffModal: closeStaffModal,
    submitStaff: submitStaff,
    addManager: addManager,
    saveNotificationPrefs: saveNotificationPrefs,
    saveAdminPin: saveAdminPin,
    sendMonthlyEmail: sendMonthlyEmail,
    toggleMonthlyView: toggleMonthlyView,
    emailMonthlyOverview: emailMonthlyOverview,
    toggleLegendPanel: toggleLegendPanel,
    loadData: loadData,
    exportCSV: exportCSV,
    switchTab: switchTab,
    goToToday: function () { state.weekOffset = 0; updateWeek(); if (state.monthlyView) loadMonthlyData(); },
    renderStaffTable: renderStaffTable,
    renderRota: renderRota,
    refreshRotaView: refreshRotaView,
    showScheduleLookup: showScheduleLookup,
    closeScheduleLookup: closeScheduleLookup,
  };

  document.addEventListener('DOMContentLoaded', init);

})();
