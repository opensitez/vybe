// vybe-test: js/with_statement_unscopables_protocol/test_js_with_statement_function_declaration_inside_non_strict
// origin: languages/js/tests/js/test_js_with_statement_unscopables_protocol.rs

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

const obj = {};
with (obj) {
    function fnInWith() { return "FuncInWith"; }
}
__check(__line(fnInWith()), "FuncInWith");
