/* AuditFlow pilot backend client: sign-in, credits, usage log, feedback, admin calls.
   Talks to a Supabase project through its REST endpoints (no extra libraries).
   Inert unless window.AF.supabase = {url, anonKey} is filled in (assets/config.js).
   Client files never go to the server: only tool name, client name typed on the form,
   line counts, minutes entered by the user and feedback text. */
(function () {
  const C = (window.AF && window.AF.supabase) || {};
  const ON = !!(C.url && C.anonKey);
  window.AFB = null;
  window.AFB_ON = ON;
  if (!ON) return;

  const $ = (s, r = document) => r.querySelector(s);
  const KEY = 'af_session_v1';
  const base = C.url.replace(/\/+$/, '');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const MSG = {
    NO_CREDITS: 'You have no credits left. Please ask your administrator for more.',
    ACCOUNT_INACTIVE: 'Your account is not active. Please contact your administrator.',
    ACCOUNT_DISABLED: 'Your account has been switched off. Please contact your administrator.',
    TOOL_UNAVAILABLE: 'This tool is switched off right now.',
    NOT_INVITED: 'This email has not been invited yet. Ask your administrator to add it, then try again.',
    'Database error saving new user': 'This email has not been invited yet. Ask your administrator to add it, then try again.',
    FORBIDDEN: 'Only administrators can do that.',
    BAD_EMAIL: 'That email address does not look right.',
    CANNOT_CHANGE_SELF: 'You cannot disable or demote your own account.',
    CANNOT_ADJUST: 'Credits cannot go below zero.',
    'Invalid login credentials': 'Email or password is not correct.',
    'User already registered': 'This email already has an account. Use Sign in instead.',
    'Email not confirmed': 'Please confirm your email first (check your inbox), then sign in.',
    EMPTY_FEEDBACK: 'Please add a rating or a comment.'
  };
  const friendly = m => { m = String(m); for (const k in MSG) if (m.includes(k)) return MSG[k]; return m; };

  let sess = null, me = null;
  try { sess = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) {}
  const save = s => { sess = s; try { s ? localStorage.setItem(KEY, JSON.stringify(s)) : localStorage.removeItem(KEY); } catch (e) {} };

  async function http(path, opt = {}) {
    const h = { apikey: C.anonKey, 'Content-Type': 'application/json' };
    // Legacy anon keys are JWTs and may be sent as Bearer; new sb_publishable_ keys must go in apikey only.
    if (opt.token) h.Authorization = 'Bearer ' + opt.token; else if (/^eyJ/.test(C.anonKey)) h.Authorization = 'Bearer ' + C.anonKey;
    let r;
    try { r = await fetch(base + path, { method: opt.method || 'GET', headers: h, body: opt.body === undefined ? undefined : JSON.stringify(opt.body) }); }
    catch (e) { const er = new Error('Could not reach the server. Check your internet connection and try again.'); er.network = true; throw er; }
    const t = await r.text(); let j = null;
    try { j = t ? JSON.parse(t) : null; } catch (e) { j = { message: t }; }
    if (!r.ok) {
      const m = (j && (j.message || j.msg || j.error_description || j.error)) || ('HTTP ' + r.status);
      const er = new Error(friendly(m)); er.raw = String(m); er.status = r.status; throw er;
    }
    return j;
  }
  const toSession = j => ({ access_token: j.access_token, refresh_token: j.refresh_token,
    expires_at: j.expires_at || Math.floor(Date.now() / 1000) + (j.expires_in || 3600), user: j.user });

  let refreshing = null;
  async function token() {
    if (!sess) return null;
    if (sess.expires_at - 60 > Date.now() / 1000) return sess.access_token;
    if (!refreshing) refreshing = http('/auth/v1/token?grant_type=refresh_token', { method: 'POST', body: { refresh_token: sess.refresh_token } })
      .then(j => { save(toSession(j)); return sess.access_token; })
      .catch(e => { if (!e.network) save(null); return null; })
      .finally(() => { refreshing = null; });
    return refreshing;
  }
  async function authed(fn) {
    const t = await token();
    if (!t) { const e = new Error('Please sign in again.'); e.signedOut = true; throw e; }
    try { return await fn(t); }
    catch (e) { if (e.status === 401) { save(null); e.signedOut = true; } throw e; }
  }
  const rpc = (name, args) => authed(t => http('/rest/v1/rpc/' + name, { method: 'POST', body: args || {}, token: t }));
  const get = path => authed(t => http('/rest/v1/' + path, { token: t }));

  const loadMe = async rec => { me = await rpc(rec ? 'record_login' : 'get_me'); return me; };

  async function signIn(email, pw) {
    const j = await http('/auth/v1/token?grant_type=password', { method: 'POST', body: { email: email.trim().toLowerCase(), password: pw } });
    save(toSession(j));
    try { await loadMe(true); } catch (e) { save(null); me = null; throw e; }
    return me;
  }
  async function signUp(email, pw, name) {
    const j = await http('/auth/v1/signup', { method: 'POST', body: { email: email.trim().toLowerCase(), password: pw, data: { full_name: (name || '').trim() } } });
    if (j && j.access_token) {
      save(toSession(j));
      try { await loadMe(true); } catch (e) { save(null); me = null; throw e; }
      return { signedIn: true };
    }
    return { signedIn: false };
  }
  async function signOut() {
    const t = sess && sess.access_token;
    if (t) { try { await http('/auth/v1/logout', { method: 'POST', token: t }); } catch (e) {} }
    save(null); me = null;
  }
  const recover = email => http('/auth/v1/recover?redirect_to=' + encodeURIComponent(location.origin + location.pathname.replace(/[^/]*$/, '') + 'login.html'),
    { method: 'POST', body: { email: email.trim().toLowerCase() } });
  const setPassword = pw => authed(t => http('/auth/v1/user', { method: 'PUT', body: { password: pw }, token: t }));
  // Password-reset links come back as login.html#access_token=...&type=recovery
  function consumeHash() {
    const h = location.hash.replace(/^#/, ''); if (!h.includes('access_token')) return null;
    const p = new URLSearchParams(h);
    save({ access_token: p.get('access_token'), refresh_token: p.get('refresh_token'),
      expires_at: +p.get('expires_at') || Math.floor(Date.now() / 1000) + (+p.get('expires_in') || 3600), user: null });
    history.replaceState(null, '', location.pathname + location.search);
    return p.get('type') || 'signup';
  }

  /* ---------- nav chip, credit note ---------- */
  const credOf = () => (me && me.profile ? me.profile.credits : 0);
  function chip() {
    const links = $('.nav .links'); if (!links) return;
    let c = $('#afchip');
    if (!c) { c = document.createElement('span'); c.id = 'afchip'; c.className = 'afchip'; links.insertBefore(c, links.querySelector('a.btn') || null); }
    if (me && me.profile) {
      const n = credOf(), cls = n <= 0 ? 'zero' : (n <= 2 ? 'low' : '');
      c.innerHTML = `<a href="account.html">${esc(me.profile.full_name || me.profile.email)}</a>${me.profile.role === 'admin' ? '' : `<span class="afcr ${cls}">${n} credit${n === 1 ? '' : 's'}</span>`}` +
        (me.profile.role === 'admin' ? '<a href="admin.html">Dashboard</a>' : '') + '<a href="#" id="afout">Sign out</a>';
      const o = $('#afout', c); if (o) o.onclick = async e => { e.preventDefault(); await signOut(); location.href = 'login.html'; };
    } else c.innerHTML = '<a href="login.html">Sign in</a>';
  }
  function rateOf(tool) { return ((me && me.rates) || []).find(r => r.tool === tool); }
  function note(tool) {
    const ph = $('.phead'); if (!ph || !me || !me.profile) return;
    let n = $('#afnote');
    if (!n) { n = document.createElement('div'); n.id = 'afnote'; n.className = 'afnote'; ph.appendChild(n); }
    const r = rateOf(tool || n.dataset.tool), c = credOf();
    if (tool) n.dataset.tool = tool;
    n.innerHTML = `Signed in as <b>${esc(me.profile.email)}</b> &middot; this tool uses <b>${r ? r.credits : 1}</b> credit per run &middot; you have <b>${c}</b> left` +
      (c <= 0 ? ' &middot; <span style="color:var(--bad)">ask your administrator for more credits</span>' : '');
  }

  /* ---------- tool pages: gate + run hooks ---------- */
  const T0 = Date.now();
  const form = $('form[data-tool]');
  if (form) document.documentElement.classList.add('af-wait');
  const page = (location.pathname.split('/').pop() || 'index.html');

  async function boot() {
    if (sess) { try { await loadMe(false); } catch (e) { if (e.signedOut || e.status === 401) save(null); } }
    chip();
    if (!form) return;
    if (!me || !me.profile) { location.replace('login.html?next=' + encodeURIComponent(page)); return; }
    if (me.profile.status !== 'active') {
      document.documentElement.classList.remove('af-wait');
      form.innerHTML = '<div class="err show" style="display:block">Your account is not active. Please contact your administrator.</div>'; return;
    }
    document.documentElement.classList.remove('af-wait');
    note(form.dataset.tool);
  }
  const ready = boot();

  const PRIV = /materiality|difference|amount|balance|value|total/i;   // never log money figures
  const AFB = {
    ready, esc, friendly, rpc, get, signIn, signUp, signOut, recover, setPassword, consumeHash,
    me: () => me, user: () => (me && me.profile) || null, session: () => sess, refresh: () => loadMe(false),
    chip, rateOf,
    async beforeRun(tool, client) {
      await ready;
      if (!me || !me.profile) throw new Error('Please sign in again.');
      const r = await rpc('start_run', { p_tool: tool, p_client: client || null });
      me.profile.credits = r.credits_left; chip(); note(tool);
      return { id: r.run_id, t0: Date.now(), tool };
    },
    async afterRun(run, ok, summary, error) {
      if (!run) return;
      try {
        const s = ok ? (summary || []).filter(x => !PRIV.test(String(x[0]))).map(x => [String(x[0]), String(x[1])]) : null;
        const r = await rpc('finish_run', { p_run: run.id, p_ok: !!ok, p_summary: s, p_error: error ? String(error).slice(0, 300) : null });
        me.profile.credits = r.credits_left; chip(); note();
      } catch (e) {}
    },
    // After a successful run: ask minutes by hand vs with the tool, rating and a comment.
    timePrompt(run, after) {
      if (!run || !after) return;
      const mins = Math.max(1, Math.round((Date.now() - T0) / 60000));
      const d = document.createElement('div'); d.className = 'aftime'; d.id = 'aftime';
      d.innerHTML = `<h4>Help us measure the time saved</h4>
        <div class="two"><div><label class="f">By hand, this would take (minutes)</label><input type="number" id="af-m" min="0" max="5000" step="1" placeholder="e.g. 120"></div>
        <div><label class="f">With the tool it took you (minutes)</label><input type="number" id="af-a" min="0" max="5000" step="1" value="${mins}"></div></div>
        <div class="stars" id="af-st" style="margin-top:10px">${[1, 2, 3, 4, 5].map(i => `<button type="button" data-v="${i}" aria-label="${i} stars">&#9733;</button>`).join('')}</div>
        <textarea id="af-c" rows="2" maxlength="2000" placeholder="Anything that worked well or needs fixing? (optional)" style="margin-top:8px"></textarea>
        <div style="display:flex;gap:10px;align-items:center;margin-top:10px"><button class="btn sm" id="af-s" type="button">Save</button><span id="af-r" class="hint"></span></div>`;
      after.insertAdjacentElement('afterend', d);
      let rating = 0; const stars = [...d.querySelectorAll('#af-st button')];
      const paint = n => stars.forEach((s, i) => s.classList.toggle('on', i < n));
      stars.forEach(s => { s.onclick = () => { rating = +s.dataset.v; paint(rating); }; });
      $('#af-s', d).onclick = async () => {
        const m = $('#af-m', d).value, a = $('#af-a', d).value, r = $('#af-r', d), b = $('#af-s', d);
        if (m === '' || a === '') { r.textContent = 'Please fill in both minutes.'; return; }
        b.disabled = true; r.textContent = 'Saving...';
        try {
          await rpc('submit_run_time', { p_run: run.id, p_manual: +m, p_actual: +a, p_rating: rating || null, p_comment: $('#af-c', d).value || null });
          d.innerHTML = '<div style="color:var(--ok);font-weight:700">Thank you - saved.</div>';
        } catch (e) { b.disabled = false; r.textContent = e.message; }
      };
    },
    async feedback(pg, cat, rating, msg) { await rpc('submit_feedback', { p_page: pg, p_category: cat, p_rating: rating || null, p_message: msg || '' }); },
    // administrator
    dashboard: () => rpc('admin_dashboard'),
    addUser: (email, name, credits, role) => rpc('admin_add_user', { p_email: email, p_name: name, p_credits: credits, p_role: role || 'user' }),
    removeInvite: email => rpc('admin_remove_invite', { p_email: email }),
    adjustCredits: (id, delta, reason) => rpc('admin_adjust_credits', { p_user: id, p_delta: delta, p_reason: reason || null }),
    setUser: (id, status, role) => rpc('admin_set_user', { p_user: id, p_status: status, p_role: role }),
    setRate: (tool, rate, credits, active) => rpc('admin_set_rate', { p_tool: tool, p_rate: rate, p_credits: credits, p_active: active }),
    setSetting: (k, v) => rpc('admin_set_setting', { p_key: k, p_value: String(v) }),
    publishFeedback: (id, on) => rpc('admin_set_feedback_published', { p_id: id, p_published: !!on }),
    xlsx, saveBlob
  };

  /* ---------- Excel (.xlsx) writer: no library, no upload. Used by the admin dashboard. ---------- */
  const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  function crc32(u8) {
    let t = crc32.t;
    if (!t) { t = crc32.t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } }
    let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = t[(c ^ u8[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function zipStore(files) {
    const enc = new TextEncoder(), parts = [], cent = []; let off = 0;
    for (const f of files) {
      const nm = enc.encode(f.name), crc = crc32(f.data), sz = f.data.length;
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(12, 0x21, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, sz, true); lh.setUint32(22, sz, true); lh.setUint16(26, nm.length, true);
      parts.push(new Uint8Array(lh.buffer), nm, f.data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(14, 0x21, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, sz, true); ch.setUint32(24, sz, true); ch.setUint16(28, nm.length, true); ch.setUint32(42, off, true);
      cent.push(new Uint8Array(ch.buffer), nm);
      off += 30 + nm.length + sz;
    }
    const csz = cent.reduce((a, b) => a + b.length, 0), end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, csz, true); end.setUint32(16, off, true);
    return new Blob(parts.concat(cent, [new Uint8Array(end.buffer)]), { type: CT });
  }
  const xe = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const colName = n => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  // Cell styles: 0 plain, 1 header, 2 bold, 3 date-time, 4 whole number, 5 title, 6 wrapped text, 7 two decimals
  const ST = { plain: 0, head: 1, bold: 2, date: 3, int: 4, title: 5, wrap: 6, dec: 7 };
  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="dd\\-mmm\\-yyyy\\ hh:mm"/></numFmts>' +
    '<fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font><font><b/><sz val="14"/><name val="Arial"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F2A44"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>' +
    '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  function sheetXml(sh) {
    const rows = sh.rows || [], hdr = !!sh.header, widths = sh.widths || [];
    let maxc = 0, body = '';
    rows.forEach((row, ri) => {
      let cells = '';
      row.forEach((c, ci) => {
        let v = c, st = null;
        if (c && typeof c === 'object' && !(c instanceof Date)) { v = c.v; st = c.s; }
        if (v == null || v === '') return;
        if (hdr && ri === 0 && st == null) st = ST.head;
        const ref = colName(ci) + (ri + 1); maxc = Math.max(maxc, ci + 1);
        if (v instanceof Date) { if (isNaN(v)) return; cells += `<c r="${ref}" s="${ST.date}"><v>${((v.getTime() - v.getTimezoneOffset() * 60000) / 86400000 + 25569).toFixed(6)}</v></c>`; }
        else if (typeof v === 'number' && isFinite(v)) cells += `<c r="${ref}"${st != null ? ` s="${st}"` : ''}><v>${v}</v></c>`;
        else cells += `<c r="${ref}"${st != null ? ` s="${st}"` : ''} t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
      });
      body += `<row r="${ri + 1}"${hdr && ri === 0 ? ' ht="30" customHeight="1"' : ''}>${cells}</row>`;
    });
    const pane = hdr ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' : '';
    const cols = widths.length ? '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>' : '';
    const af = hdr && rows.length > 1 && maxc ? `<autoFilter ref="A1:${colName(maxc - 1)}${rows.length}"/>` : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="' + (sh.grid === false ? 0 : 1) + '">' + pane + '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="14"/>' + cols + '<sheetData>' + body + '</sheetData>' + af + '</worksheet>';
  }
  // sheets: [{name, rows: [[cell,...]], header: bool, widths: [..], grid: bool}]  ->  Blob (.xlsx)
  function xlsx(sheets) {
    const enc = new TextEncoder(), X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', files = [];
    const names = sheets.map((s, i) => (String(s.name || 'Sheet' + (i + 1)).replace(/[\[\]:*?\/\\]/g, ' ').slice(0, 31)) || 'Sheet' + (i + 1));
    const add = (name, text) => files.push({ name, data: enc.encode(text) });
    add('[Content_Types].xml', X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>');
    add('_rels/.rels', X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    add('xl/workbook.xml', X + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + names.map((n, i) => `<sheet name="${xe(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>');
    add('xl/_rels/workbook.xml.rels', X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    add('xl/styles.xml', STYLES);
    sheets.forEach((s, i) => add(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s)));
    return zipStore(files);
  }
  function saveBlob(name, blob) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  /* ---------- published feedback shown on the home page ---------- */
  async function testimonials() {
    const how = document.getElementById('how'); if (!how) return;
    let list; try { list = await http('/rest/v1/rpc/public_testimonials', { method: 'POST', body: {} }); } catch (e) { return; }
    if (!Array.isArray(list) || !list.length) return;
    const sec = document.createElement('section'); sec.id = 'voices';
    sec.innerHTML = '<div class="wrap"><div class="center"><span class="pill">What users say</span><h2 style="margin-top:14px">Feedback from <span class="grad">the team</span></h2></div><div class="feats aftest">' +
      list.map(t => `<figure class="feat"><div class="aftst">${t.rating ? '&#9733;'.repeat(t.rating) + '<span style="opacity:.25">' + '&#9733;'.repeat(5 - t.rating) + '</span>' : ''}</div><blockquote>${esc(t.message)}</blockquote><figcaption><b>${esc(t.name)}</b>${t.tool ? ' &middot; ' + esc(t.tool) : ''}</figcaption></figure>`).join('') + '</div></div>';
    const host = how.closest('section') || how; host.parentNode.insertBefore(sec, host.nextSibling);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', testimonials); else testimonials();
  window.AFB = AFB;
})();
