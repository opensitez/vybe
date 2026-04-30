// PHP string helpers — JS-source polyfills bundled as `__vybe_php_*`.
// Compose from JS String.prototype.* primitives so PHP gets the
// surface without PHP-specific Rust host fns.

// ucwords(str, delimiters?) — uppercase the first letter of each word.
// PHP defaults delimiters to whitespace + form-feed + vertical-tab.
// Compare via charCode to avoid Vybe's JS string parser not honouring
// `\f` / `\v` escape sequences (it keeps them as literal "\\f" pairs).
function __vybe_php_ucwords(str, delims) {
    var s = "" + str;
    var out = "";
    var capitalize = true;
    for (var i = 0; i < s.length; i++) {
        var code = s.charCodeAt(i);
        var c = s.charAt(i);
        var isDelim;
        if (delims === undefined || delims === null) {
            // Default: any whitespace per PHP's `php_charmask`.
            isDelim = code === 32 || code === 9 || code === 10 || code === 11 || code === 12 || code === 13;
        } else {
            isDelim = delims.indexOf(c) >= 0;
        }
        if (isDelim) {
            out += c;
            capitalize = true;
        } else if (capitalize) {
            out += c.toUpperCase();
            capitalize = false;
        } else {
            out += c;
        }
    }
    return out;
}

// str_split(str, length?) — split into chunks. Default length 1
// returns array of single chars; bigger length returns chunks.
function __vybe_php_str_split(str, length) {
    if (length === undefined || length === null) length = 1;
    if (length < 1) return false;
    var out = [];
    for (var i = 0; i < str.length; i += length) {
        out.push(str.substring(i, i + length));
    }
    return out;
}

// str_pad(str, length, padStr?, padType?) — pad to `length` with
// `padStr` (default " ") on `padType` side (0=left, 1=right [default],
// 2=both). PHP STR_PAD_* constants are 0/1/2 respectively.
function __vybe_php_str_pad(str, length, padStr, padType) {
    str = "" + str;
    if (padStr === undefined || padStr === null || padStr === "") padStr = " ";
    if (padType === undefined || padType === null) padType = 1;
    if (str.length >= length) return str;
    var diff = length - str.length;
    var pad = "";
    while (pad.length < diff) pad += padStr;
    pad = pad.substring(0, diff);
    if (padType === 0) return pad + str;
    if (padType === 2) {
        var leftLen = Math.floor(diff / 2);
        var rightLen = diff - leftLen;
        var leftPad = "";
        while (leftPad.length < leftLen) leftPad += padStr;
        leftPad = leftPad.substring(0, leftLen);
        var rightPad = "";
        while (rightPad.length < rightLen) rightPad += padStr;
        rightPad = rightPad.substring(0, rightLen);
        return leftPad + str + rightPad;
    }
    return str + pad;
}

// substr_count(haystack, needle, offset?, length?) — count
// non-overlapping occurrences of needle in haystack.
function __vybe_php_substr_count(hay, needle, offset, length) {
    if (offset === undefined || offset === null) offset = 0;
    var slice;
    if (length === undefined || length === null) {
        slice = hay.substring(offset);
    } else {
        slice = hay.substring(offset, offset + length);
    }
    if (needle.length === 0) return 0;
    var count = 0;
    var pos = 0;
    while (true) {
        var idx = slice.indexOf(needle, pos);
        if (idx < 0) break;
        count++;
        pos = idx + needle.length;
    }
    return count;
}

// substr_replace(str, replacement, start, length?) — replace a slice.
function __vybe_php_substr_replace(str, repl, start, length) {
    var len = str.length;
    var s = start < 0 ? Math.max(len + start, 0) : Math.min(start, len);
    var l;
    if (length === undefined || length === null) {
        l = len - s;
    } else if (length < 0) {
        l = Math.max(len + length - s, 0);
    } else {
        l = Math.min(length, len - s);
    }
    return str.substring(0, s) + repl + str.substring(s + l);
}

// str_ireplace(search, replace, subject) — case-insensitive replace.
function __vybe_php_str_ireplace(search, repl, subj) {
    var s = "" + subj;
    var srch = "" + search;
    if (srch.length === 0) return s;
    // Walk the source by chunks delimited by case-insensitive matches
    // of `srch`. Each chunk before a match goes through verbatim;
    // matches are replaced with `repl`. Recursive `slice` (rather than
    // an in-place `lower.indexOf(srch, pos)` accumulator) avoids a
    // compile-time interaction in the polyfill cache where the loop
    // variant produced wrong results when invoked across language
    // contexts.
    var srchLower = srch.toLowerCase();
    var srchLen = srch.length;
    var parts = [];
    var rest = s;
    while (rest.length > 0) {
        var i = rest.toLowerCase().indexOf(srchLower);
        if (i < 0) {
            parts.push(rest);
            break;
        }
        parts.push(rest.slice(0, i));
        parts.push(repl);
        rest = rest.slice(i + srchLen);
    }
    return parts.join("");
}

// str_word_count(str) — count whitespace-separated words.
function __vybe_php_str_word_count(str) {
    var parts = ("" + str).split(/[\s,.!?;:]+/);
    var count = 0;
    for (var i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) count++;
    }
    return count;
}

// strstr(haystack, needle, before?) — substring from first match (or
// before it if `before` is true).
function __vybe_php_strstr(hay, needle, before) {
    var idx = ("" + hay).indexOf("" + needle);
    if (idx < 0) return false;
    if (before === true) return hay.substring(0, idx);
    return hay.substring(idx);
}

function __vybe_php_stristr(hay, needle, before) {
    var lower = ("" + hay).toLowerCase();
    var n = ("" + needle).toLowerCase();
    var idx = lower.indexOf(n);
    if (idx < 0) return false;
    if (before === true) return hay.substring(0, idx);
    return hay.substring(idx);
}

// urlencode(str) — PHP's percent-encoding (space → "+"). Manual hex
// because `(n).toString(16)` doesn't honour the radix arg.
function __vybe_php_urlencode(str) {
    var s = "" + str;
    var out = "";
    var safe = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~";
    var HEX = "0123456789ABCDEF";
    for (var i = 0; i < s.length; i++) {
        var code = s.charCodeAt(i);
        if (code === 32) {
            out += "+";
        } else {
            var c = s.charAt(i);
            if (safe.indexOf(c) >= 0) {
                out += c;
            } else {
                out += "%" + HEX.charAt((code >> 4) & 0xF) + HEX.charAt(code & 0xF);
            }
        }
    }
    return out;
}

// rawurlencode(str) — RFC 3986; space → "%20".
function __vybe_php_rawurlencode(str) {
    return __vybe_php_urlencode(str).split("+").join("%20");
}

// urldecode(str) — reverse of urlencode.
function __vybe_php_urldecode(str) {
    return decodeURIComponent(("" + str).split("+").join(" "));
}

// bin2hex(str) — convert each byte to two hex digits. Manual hex
// because `Number.prototype.toString(radix)` doesn't honour the
// radix arg in Vybe's JS profile (always returns decimal).
function __vybe_php_bin2hex(str) {
    var s = "" + str;
    var hex_digits = "0123456789abcdef";
    var out = "";
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        out += hex_digits.charAt((c >> 4) & 0xF);
        out += hex_digits.charAt(c & 0xF);
    }
    return out;
}

// hex2bin(hex) — reverse of bin2hex. Manual parse because `parseInt`
// with radix may also drop the radix arg.
function __vybe_php_hex2bin(hex) {
    var s = ("" + hex).toLowerCase();
    var hex_digits = "0123456789abcdef";
    var out = "";
    for (var i = 0; i + 1 < s.length; i += 2) {
        var hi = hex_digits.indexOf(s.charAt(i));
        var lo = hex_digits.indexOf(s.charAt(i + 1));
        if (hi < 0 || lo < 0) return false;
        out += String.fromCharCode((hi << 4) | lo);
    }
    return out;
}

// chunk_split(str, length?, end?) — split into chunks, each followed
// by `end`. PHP defaults length=76, end="\r\n".
function __vybe_php_chunk_split(str, length, end) {
    if (length === undefined || length === null) length = 76;
    if (end === undefined || end === null) end = "\r\n";
    if (length < 1) return false;
    var s = "" + str;
    var out = "";
    for (var i = 0; i < s.length; i += length) {
        out += s.substring(i, i + length) + end;
    }
    return out;
}

// wordwrap(str, width?, break_str?, cut?) — wrap to `width` chars,
// breaking on whitespace (or mid-word if cut=true).
function __vybe_php_wordwrap(str, width, br, cut) {
    if (width === undefined || width === null) width = 75;
    if (br === undefined || br === null) br = "\n";
    if (cut === undefined || cut === null) cut = false;
    var s = "" + str;
    var lines = s.split("\n");
    var out = [];
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        if (line.length <= width) {
            out.push(line);
            continue;
        }
        var words = line.split(" ");
        var current = "";
        for (var j = 0; j < words.length; j++) {
            var word = words[j];
            if (current.length === 0) {
                current = word;
            } else if (current.length + 1 + word.length <= width) {
                current += " " + word;
            } else {
                out.push(current);
                current = word;
            }
            if (cut && current.length > width) {
                while (current.length > width) {
                    out.push(current.substring(0, width));
                    current = current.substring(width);
                }
            }
        }
        if (current.length > 0) out.push(current);
    }
    return out.join(br);
}

// str_replace(search, replace, subject) — PHP semantics: search and
// replace can each be a string or array. The opcode `STR_REPLACE`
// only handles string-string-string; this polyfill covers the array
// shapes (PHP §str_replace).
function __vybe_php_str_replace(search, repl, subj) {
    var s = "" + subj;
    var searchArr = Array.isArray(search) ? search : [search];
    var replArr = Array.isArray(repl) ? repl : null;
    for (var i = 0; i < searchArr.length; i++) {
        var needle = "" + searchArr[i];
        if (needle.length === 0) continue;
        var rep;
        if (replArr === null) {
            rep = "" + repl;
        } else if (i < replArr.length) {
            rep = "" + replArr[i];
        } else {
            rep = "";
        }
        var pieces = s.split(needle);
        s = pieces.join(rep);
    }
    return s;
}

// ctype_* — character-class predicates. PHP returns false on empty
// input; otherwise true iff every char matches the class.
function __vybe_php_ctype_alpha(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        var ok = (c >= 65 && c <= 90) || (c >= 97 && c <= 122);
        if (!ok) return false;
    }
    return true;
}
function __vybe_php_ctype_digit(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        if (c < 48 || c > 57) return false;
    }
    return true;
}
function __vybe_php_ctype_alnum(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        var ok = (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122);
        if (!ok) return false;
    }
    return true;
}
function __vybe_php_ctype_space(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        var ok = c === 32 || c === 9 || c === 10 || c === 11 || c === 12 || c === 13;
        if (!ok) return false;
    }
    return true;
}
function __vybe_php_ctype_upper(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        if (c < 65 || c > 90) return false;
    }
    return true;
}
function __vybe_php_ctype_lower(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        if (c < 97 || c > 122) return false;
    }
    return true;
}
function __vybe_php_ctype_xdigit(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        var ok = (c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102);
        if (!ok) return false;
    }
    return true;
}
function __vybe_php_ctype_punct(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        var ok = (c >= 33 && c <= 47) || (c >= 58 && c <= 64) || (c >= 91 && c <= 96) || (c >= 123 && c <= 126);
        if (!ok) return false;
    }
    return true;
}
function __vybe_php_ctype_print(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        if (c < 32 || c > 126) return false;
    }
    return true;
}
function __vybe_php_ctype_cntrl(s) {
    s = "" + s;
    if (s.length === 0) return false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        var ok = (c >= 0 && c <= 31) || c === 127;
        if (!ok) return false;
    }
    return true;
}

// number_format(num, decimals?, decsep?, thousep?) — PHP default
// formatting: "1,234.56" with decimals=2. Other separators allow
// locale-style overrides.
function __vybe_php_number_format(num, decimals, decsep, thousep) {
    if (decimals === undefined || decimals === null) decimals = 0;
    if (decsep === undefined || decsep === null) decsep = ".";
    if (thousep === undefined || thousep === null) thousep = ",";
    var n = +num;
    if (n !== n) return "0";
    // Take abs via condition rather than ternary on a Bool var that
    // gets shadowed by later compiler-emitted scratch slots.
    var sign = "";
    if (n < 0) { sign = "-"; n = -n; }
    // PHP rounds half away from zero (`number_format(0.5, 0)` → "1");
    // JS `toFixed` follows the spec's banker's rounding in most
    // engines (`(0.5).toFixed(0)` → "0"). Pre-round through Math.round
    // — Vybe's Math.round is half-away-from-zero, matching PHP.
    var scale = Math.pow(10, decimals);
    n = Math.round(n * scale) / scale;
    var fixed = n.toFixed(decimals);
    var parts = fixed.split(".");
    var intPart = parts[0];
    var fracPart = parts.length > 1 ? parts[1] : "";
    var withSep = "";
    var len = intPart.length;
    for (var i = 0; i < len; i++) {
        if (i > 0 && (len - i) % 3 === 0) withSep += thousep;
        withSep += intPart.charAt(i);
    }
    var out = withSep;
    if (fracPart.length > 0) out += decsep + fracPart;
    return sign + out;
}
