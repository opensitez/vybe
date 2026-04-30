// PHP math helpers — JS-source polyfills.
//
// Composes from JS Math primitives so PHP gets the surface without
// PHP-specific Rust host fns. Variadic min/max + base conversion +
// dechex/decbin/decoct round out the math_advanced surface.

// PHP min/max accept (a, b, ...args) OR (array). The polyfill detects
// the single-array form and iterates; otherwise iterates the args.
function __vybe_php_min(...args) {
    var values;
    if (args.length === 1 && Array.isArray(args[0])) {
        values = args[0];
    } else {
        values = args;
    }
    if (values.length === 0) return false;
    var best = values[0];
    for (var j = 1; j < values.length; j++) {
        if (+values[j] < +best) best = values[j];
    }
    return best;
}

function __vybe_php_max(...args) {
    var values;
    if (args.length === 1 && Array.isArray(args[0])) {
        values = args[0];
    } else {
        values = args;
    }
    if (values.length === 0) return false;
    var best = values[0];
    for (var j = 1; j < values.length; j++) {
        if (+values[j] > +best) best = values[j];
    }
    return best;
}

// dechex / decbin / decoct — convert decimal to hex/bin/oct string.
// Manual base conversion since `Number.prototype.toString(radix)` may
// not honour the radix arg.
function __vybe_php_decbin(n) {
    n = +n;
    if (n === 0) return "0";
    if (n < 0) n = n >>> 0; // PHP wraps to unsigned 32-bit for negatives.
    var out = "";
    while (n > 0) {
        out = (n & 1) + out;
        n = Math.floor(n / 2);
    }
    return out;
}

function __vybe_php_decoct(n) {
    n = +n;
    if (n === 0) return "0";
    if (n < 0) n = n >>> 0;
    var out = "";
    while (n > 0) {
        out = (n & 7) + out;
        n = Math.floor(n / 8);
    }
    return out;
}

function __vybe_php_dechex(n) {
    n = +n;
    if (n === 0) return "0";
    if (n < 0) n = n >>> 0;
    var hex = "0123456789abcdef";
    var out = "";
    while (n > 0) {
        out = hex.charAt(n & 0xF) + out;
        n = Math.floor(n / 16);
    }
    return out;
}

// base_convert(num, frombase, tobase) — string-to-string conversion.
function __vybe_php_base_convert(num, from, to) {
    var s = ("" + num).toLowerCase();
    var digits = "0123456789abcdefghijklmnopqrstuvwxyz";
    var n = 0;
    for (var i = 0; i < s.length; i++) {
        var d = digits.indexOf(s.charAt(i));
        if (d < 0 || d >= from) continue;
        n = n * from + d;
    }
    if (n === 0) return "0";
    var out = "";
    while (n > 0) {
        out = digits.charAt(n % to) + out;
        n = Math.floor(n / to);
    }
    return out;
}
