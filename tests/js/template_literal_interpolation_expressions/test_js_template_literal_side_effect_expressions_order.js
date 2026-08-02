// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_side_effect_expressions_order
// origin: languages/js/tests/js/test_js_template_literal_interpolation_expressions.rs

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

let counter = 0;
function inc() { return ++counter; }
__check(__line(`${inc()}-${inc()}-${inc()}`), "1-2-3");
