// PHP array_* helpers — JS-source polyfills bundled as `__vybe_php_*`.
//
// Only the multi-step / loop-heavy operations live here. Single-op
// PHP functions (sort/rsort/usort/array_unique/array_fill/
// array_product) are already routed via the PHP profile through
// existing common emits or `__vybe_*` stdlib chunks so they don't
// need PHP-specific bytecode.
//
// Every helper is written in terms of `Array.prototype.*` and
// `Object.*` — the JS compiler resolves these to the same
// `ecma:array.*` / `ecma:object.*` host fns the JS profile uses,
// so a single host implementation backs every language.

// array_pad(arr, size, value) — abs(size) target length; negative pads left.
function __vybe_php_array_pad(arr, size, value) {
    var len = arr.length;
    var target = size < 0 ? -size : size;
    if (target <= len) return arr.slice();
    var diff = target - len;
    var pad = [];
    for (var i = 0; i < diff; i++) pad.push(value);
    if (size < 0) return pad.concat(arr);
    return arr.slice().concat(pad);
}

// array_chunk(arr, size) — partition into equal-sized chunks.
function __vybe_php_array_chunk(arr, size) {
    var out = [];
    if (size < 1) return out;
    for (var i = 0; i < arr.length; i += size) {
        out.push(arr.slice(i, i + size));
    }
    return out;
}

// array_flip(obj) — swap keys and values.
function __vybe_php_array_flip(obj) {
    var out = {};
    if (Array.isArray(obj)) {
        for (var i = 0; i < obj.length; i++) out[obj[i]] = i;
    } else {
        var keys = Object.keys(obj);
        for (var j = 0; j < keys.length; j++) out[obj[keys[j]]] = keys[j];
    }
    return out;
}

// array_combine(keys, values) — zip into an assoc array.
function __vybe_php_array_combine(keys, values) {
    var out = {};
    for (var i = 0; i < keys.length; i++) out[keys[i]] = values[i];
    return out;
}

// array_diff(a, b) — values in a not in b.
function __vybe_php_array_diff(a, b) {
    var seen = {};
    for (var i = 0; i < b.length; i++) seen["" + b[i]] = true;
    var out = [];
    for (var j = 0; j < a.length; j++) if (!seen["" + a[j]]) out.push(a[j]);
    return out;
}

// array_intersect(a, b) — values present in both.
function __vybe_php_array_intersect(a, b) {
    var seen = {};
    for (var i = 0; i < b.length; i++) seen["" + b[i]] = true;
    var out = [];
    for (var j = 0; j < a.length; j++) if (seen["" + a[j]]) out.push(a[j]);
    return out;
}

// array_diff_assoc(a, b) — entries in a whose key→value pair differs in b.
function __vybe_php_array_diff_assoc(a, b) {
    var out = {};
    var keys = Object.keys(a);
    for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (!(k in b) || ("" + b[k]) !== ("" + a[k])) out[k] = a[k];
    }
    return out;
}

// array_intersect_key(a, b) — entries from a whose keys exist in b.
function __vybe_php_array_intersect_key(a, b) {
    var out = {};
    var keys = Object.keys(a);
    for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (k in b) out[k] = a[k];
    }
    return out;
}

// array_replace(a, b) — a with b entries overwriting matching keys.
function __vybe_php_array_replace(a, b) {
    var out = {};
    var ak = Object.keys(a);
    for (var i = 0; i < ak.length; i++) out[ak[i]] = a[ak[i]];
    var bk = Object.keys(b);
    for (var j = 0; j < bk.length; j++) out[bk[j]] = b[bk[j]];
    return out;
}

// array_count_values(arr) — frequency map.
function __vybe_php_array_count_values(arr) {
    var out = {};
    for (var i = 0; i < arr.length; i++) {
        var k = "" + arr[i];
        out[k] = (out[k] || 0) + 1;
    }
    return out;
}

// array_column(rows, col, indexKey?) — pluck a column across rows.
function __vybe_php_array_column(rows, col, indexKey) {
    var keyed = indexKey !== undefined && indexKey !== null;
    if (keyed) {
        var out = {};
        for (var i = 0; i < rows.length; i++) out[rows[i][indexKey]] = rows[i][col];
        return out;
    }
    var arr = [];
    for (var j = 0; j < rows.length; j++) arr.push(rows[j][col]);
    return arr;
}

// array_key_first(arr) — first key (number for arrays, string for objects).
function __vybe_php_array_key_first(obj) {
    if (Array.isArray(obj)) return obj.length === 0 ? null : 0;
    var keys = Object.keys(obj);
    return keys.length === 0 ? null : keys[0];
}

// array_key_last(arr) — last key.
function __vybe_php_array_key_last(obj) {
    if (Array.isArray(obj)) return obj.length === 0 ? null : obj.length - 1;
    var keys = Object.keys(obj);
    return keys.length === 0 ? null : keys[keys.length - 1];
}

// Assoc-sort family (asort/arsort/ksort/krsort/uasort/uksort).
//
// PHP associative arrays are JS Maps in Vybe's model, so these
// polyfills use Map.prototype.{get, set, delete, clear, keys}
// instead of property access — `delete obj[k]` and `obj[k] = v`
// don't reach the Map backing storage. Mutating in place keeps
// PHP-by-reference semantics working since the caller's variable
// points to the same Map.
//
// Bubble sort instead of `Array.prototype.sort(cb)` because
// `build_polyfill` only extracts the named export chunk and drops
// the inner comparator closures.

function __vybe_php_asort(obj) {
    var keys = Object.keys(obj);
    var n = keys.length;
    for (var i = 0; i < n; i++) {
        for (var j = 0; j < n - 1 - i; j++) {
            var va = +obj.get(keys[j]), vb = +obj.get(keys[j + 1]);
            if (va > vb) {
                var tmp = keys[j];
                keys[j] = keys[j + 1];
                keys[j + 1] = tmp;
            }
        }
    }
    var entries = [];
    for (var p = 0; p < n; p++) entries.push([keys[p], obj.get(keys[p])]);
    obj.clear();
    for (var r = 0; r < n; r++) obj.set(entries[r][0], entries[r][1]);
}

function __vybe_php_arsort(obj) {
    var keys = Object.keys(obj);
    var n = keys.length;
    for (var i = 0; i < n; i++) {
        for (var j = 0; j < n - 1 - i; j++) {
            var va = +obj.get(keys[j]), vb = +obj.get(keys[j + 1]);
            if (va < vb) {
                var tmp = keys[j];
                keys[j] = keys[j + 1];
                keys[j + 1] = tmp;
            }
        }
    }
    var entries = [];
    for (var p = 0; p < n; p++) entries.push([keys[p], obj.get(keys[p])]);
    obj.clear();
    for (var r = 0; r < n; r++) obj.set(entries[r][0], entries[r][1]);
}

function __vybe_php_ksort(obj) {
    var keys = Object.keys(obj);
    keys.sort();
    var entries = [];
    for (var i = 0; i < keys.length; i++) entries.push([keys[i], obj.get(keys[i])]);
    obj.clear();
    for (var k = 0; k < entries.length; k++) obj.set(entries[k][0], entries[k][1]);
}

function __vybe_php_krsort(obj) {
    var keys = Object.keys(obj);
    keys.sort();
    keys.reverse();
    var entries = [];
    for (var i = 0; i < keys.length; i++) entries.push([keys[i], obj.get(keys[i])]);
    obj.clear();
    for (var k = 0; k < entries.length; k++) obj.set(entries[k][0], entries[k][1]);
}

function __vybe_php_uasort(obj, cmp) {
    var keys = Object.keys(obj);
    var n = keys.length;
    for (var i = 0; i < n; i++) {
        for (var j = 0; j < n - 1 - i; j++) {
            if (cmp(obj.get(keys[j]), obj.get(keys[j + 1])) > 0) {
                var tmp = keys[j];
                keys[j] = keys[j + 1];
                keys[j + 1] = tmp;
            }
        }
    }
    var entries = [];
    for (var p = 0; p < n; p++) entries.push([keys[p], obj.get(keys[p])]);
    obj.clear();
    for (var r = 0; r < n; r++) obj.set(entries[r][0], entries[r][1]);
}

function __vybe_php_uksort(obj, cmp) {
    var keys = Object.keys(obj);
    var n = keys.length;
    for (var i = 0; i < n; i++) {
        for (var j = 0; j < n - 1 - i; j++) {
            if (cmp(keys[j], keys[j + 1]) > 0) {
                var tmp = keys[j];
                keys[j] = keys[j + 1];
                keys[j + 1] = tmp;
            }
        }
    }
    var entries = [];
    for (var p = 0; p < n; p++) entries.push([keys[p], obj.get(keys[p])]);
    obj.clear();
    for (var r = 0; r < n; r++) obj.set(entries[r][0], entries[r][1]);
}
