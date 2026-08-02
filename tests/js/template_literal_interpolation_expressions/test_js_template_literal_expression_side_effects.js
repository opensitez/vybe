// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_expression_side_effects
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

const trace = [];
const payload = {
    get value() {
        trace.push("read");
        return 7;
    }
};
__check(__line(`${payload.value}:${payload.value}`), "7:7");
__check(__line(trace.length), "2");
