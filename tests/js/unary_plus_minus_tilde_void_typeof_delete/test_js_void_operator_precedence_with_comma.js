// vybe-test: js/unary_plus_minus_tilde_void_typeof_delete/test_js_void_operator_precedence_with_comma
// origin: languages/js/tests/js/test_js_unary_plus_minus_tilde_void_typeof_delete.rs

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

let x = 0;
const res = void (x += 1, x += 10);
__check(__line(`${res === undefined}|${x}`), "true|11");
