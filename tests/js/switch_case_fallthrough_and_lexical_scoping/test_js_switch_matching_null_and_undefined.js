// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_matching_null_and_undefined
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

let res = "";
switch(null) {
    case undefined: res = "undef"; break;
    case null: res = "nullVal"; break;
}
__check(__line(res), "nullVal");
