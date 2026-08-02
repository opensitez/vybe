// vybe-test: js/regexp_v_flag_set_notation_intersection_subtraction/test_js_regexp_v_flag_mutually_exclusive_with_u_flag_throws_syntaxerror
// origin: languages/js/tests/js/test_js_regexp_v_flag_set_notation_intersection_subtraction.rs

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

try {
    eval("const re = /test/uv;"); // Flags u and v cannot be combined!
} catch (e) {
    __check(__line("u and v flags SyntaxError"), "u and v flags SyntaxError");
}
