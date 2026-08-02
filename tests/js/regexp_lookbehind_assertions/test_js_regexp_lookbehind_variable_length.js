// vybe-test: js/regexp_lookbehind_assertions/test_js_regexp_lookbehind_variable_length
// origin: languages/js/tests/js/test_js_regexp_lookbehind_assertions.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const re = /(?<=(a|bb))\d+/; // JS supports variable-length lookbehind!
const m1 = re.exec("a123");
const m2 = re.exec("bb456");
__check(__line(m1[0] + "|" + m2[0]), "123|456");
