// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_assignment_expression
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

let val;
__check(__line(`Assigned: ${val = 100}`), "Assigned: 100");
__check(__line(val), "100");
