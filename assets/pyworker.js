/* Runs the Python audit tools (openpyxl etc.) inside the browser via Pyodide. Nothing is uploaded anywhere. */
import { loadPyodide } from '../pyodide/pyodide.mjs';
const BASE = new URL('../', self.location.href);
let py = null, ready = null;

function init() {
  if (ready) return ready;
  ready = (async () => {
    const man = await (await fetch(new URL('manifest.json', BASE))).json();
    py = await loadPyodide({ indexURL: new URL('pyodide/', BASE).href });
    for (const w of man.wheels) {
      const buf = await (await fetch(new URL('wheels/' + w, BASE))).arrayBuffer();
      py.unpackArchive(buf, 'wheel');
    }
    py.FS.mkdirTree('/app');
    for (const f of man.py) {
      py.FS.writeFile('/app/' + f, await (await fetch(new URL('py/' + f, BASE))).text());
    }
    py.runPython("import sys; sys.path.insert(0, '/app'); import bridge");
    postMessage({ type: 'ready' });
  })().catch(e => { postMessage({ type: 'fatal', error: String(e) }); throw e; });
  return ready;
}

function cleanError(e) {
  const lines = String(e && e.message ? e.message : e).trim().split('\n').filter(Boolean);
  let last = lines[lines.length - 1] || 'Something went wrong.';
  return last.replace(/^(builtins\.)?ValueError:\s*/, '');
}

self.onmessage = async (ev) => {
  const d = ev.data;
  if (d.cmd === 'init') { try { await init(); } catch (e) {} return; }
  if (d.cmd === 'template') {
    try {
      await init();
      py.globals.set('_M', d.module);
      py.runPython("import importlib; importlib.import_module(_M).make_template('/tmp/_tpl.xlsx')");
      const bytes = py.FS.readFile('/tmp/_tpl.xlsx');
      postMessage({ id: d.id, bytes: bytes.buffer }, [bytes.buffer]);
    } catch (e) { postMessage({ id: d.id, error: cleanError(e) }); }
    return;
  }
  if (d.cmd !== 'run') return;
  try {
    await init();
    py.globals.set('_T', d.tool);
    py.globals.set('_F', py.toPy(d.fields));
    py.globals.set('_N', d.filename);
    py.globals.set('_B', new Uint8Array(d.buf));
    const ex = {};
    for (const [k, v] of Object.entries(d.extras || {})) {
      py.globals.set('_X_' + k, new Uint8Array(v.buf));
      ex[k] = v.name;
    }
    py.globals.set('_XN', py.toPy(ex));
    py.runPython(`
_extras = {k: (n, globals()['_X_' + k].to_py().tobytes()) for k, n in dict(_XN).items()}
_name, _data, _summary = bridge.run(_T, dict(_F), _N, _B.to_py().tobytes(), _extras)
open('/tmp/_out.xlsx', 'wb').write(_data)
_summary_json = __import__('json').dumps([list(map(str, x)) for x in _summary])
`);
    const bytes = py.FS.readFile('/tmp/_out.xlsx');
    const name = py.globals.get('_name');
    const summary = JSON.parse(py.globals.get('_summary_json'));
    postMessage({ id: d.id, name, summary, bytes: bytes.buffer }, [bytes.buffer]);
  } catch (e) {
    postMessage({ id: d.id, error: cleanError(e) });
  }
};
