// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_symbol_matching
// origin: languages/js/tests/js/test_js_switch_case_fallthrough_and_lexical_scoping.rs

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

const s1 = Symbol("a");
const s2 = Symbol("b");
let res = "";
switch(s1) {
    case s2: res = "s2"; break;
    case s1: res = "s1"; break;
}
__check(__line(res), "s1");
