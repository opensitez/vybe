const fs = require('fs');
const path = require('path');
const vm = require('vm');
// no external deps; use recursive fs walk

function convertFile(file) {
  const code = fs.readFileSync(file, 'utf8');
  const sandbox = {
    moduleResult: null,
    require: () => ({}),
    console: console,
    // Provide minimal monaco stub used by some language files
    monaco: {
      languages: {
        IndentAction: { None: 0, Indent: 1, IndentOutdent: 2 }
      }
    }
  };

  sandbox.define = function() {
    // define(name?, deps?, factory)
    const args = Array.prototype.slice.call(arguments);
    const factory = args[args.length - 1];
    try {
      sandbox.moduleResult = factory(function() { return {}; });
    } catch (e) {
      throw e;
    }
  };

  try {
    vm.createContext(sandbox);
    vm.runInContext(code, sandbox, { filename: file });
  } catch (e) {
    console.error('Failed to eval', file, e && e.message);
    return false;
  }

  const result = sandbox.moduleResult;
  if (!result || (!result.conf && !result.language)) {
    console.warn('No conf/language found in', file);
    return false;
  }

  function replacer(key, value) {
    if (value instanceof RegExp) return value.toString();
    if (typeof value === 'function') return value.toString();
    return value;
  }

  const outObj = {
    conf: result.conf || null,
    language: result.language || null
  };

  // Normalize RegExp and function values into strings so they serialize cleanly
  function normalize(v) {
    if (v instanceof RegExp) return v.toString();
    if (typeof v === 'function') return v.toString();
    if (Array.isArray(v)) return v.map(normalize);
    if (v && typeof v === 'object') {
      const o = {};
      for (const k of Object.keys(v)) o[k] = normalize(v[k]);
      return o;
    }
    return v;
  }

  const normalized = normalize(outObj);

  // Custom pretty serializer that groups string arrays into multiple items per line
  function serialize(val, level = 0) {
    const indent = '  '.repeat(level);
    if (val === null) return 'null';
    if (Array.isArray(val)) {
      if (val.length === 0) return '[]';
      const allStrings = val.every(x => typeof x === 'string');
      if (allStrings && val.length > 1) {
        const itemsPerLine = 6;
        let out = '[\n';
        for (let i = 0; i < val.length; i += itemsPerLine) {
          const chunk = val.slice(i, i + itemsPerLine).map(s => JSON.stringify(s)).join(', ');
          out += indent + '  ' + chunk + (i + itemsPerLine < val.length ? ',\n' : '\n');
        }
        out += indent + ']';
        return out;
      }
      let out = '[\n';
      for (let i = 0; i < val.length; i++) {
        out += indent + '  ' + serialize(val[i], level + 1) + (i < val.length - 1 ? ',\n' : '\n');
      }
      out += indent + ']';
      return out;
    }
    if (val && typeof val === 'object') {
      const keys = Object.keys(val);
      if (keys.length === 0) return '{}';
      let out = '{\n';
      for (let i = 0; i < keys.length; i++) {
        const k = keys[i];
        out += indent + '  ' + JSON.stringify(k) + ': ' + serialize(val[k], level + 1) + (i < keys.length - 1 ? ',\n' : '\n');
      }
      out += indent + '}';
      return out;
    }
    return JSON.stringify(val);
  }

  const outPath = file.replace(/\.js$/, '.json');
  fs.writeFileSync(outPath, serialize(normalized, 0), 'utf8');
  console.log('Wrote', outPath);
  return true;
}

function main() {
  const base = path.join(__dirname, '..', 'basic-languages');
  const files = [];
  function walk(dir) {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) walk(full);
      else if (e.isFile() && full.endsWith('.js')) files.push(full);
    }
  }
  walk(base);
  console.log('Found', files.length, '.js files under basic-languages');
  let success = 0;
  for (const f of files) {
    try {
      if (convertFile(f)) success++;
    } catch (e) {
      console.error('Error converting', f, e && e.message);
    }
  }
  console.log(`Converted ${success}/${files.length}`);
}

main();
