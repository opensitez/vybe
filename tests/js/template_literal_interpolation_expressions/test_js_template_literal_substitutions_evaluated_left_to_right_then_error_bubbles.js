// vybe-test: js/template_literal_interpolation_expressions/test_js_template_literal_substitutions_evaluated_left_to_right_then_error_bubbles
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
function boom() {
    trace.push("boom");
    throw new Error("interpolation-failed");
}
try {
    console.log(`a=${(() => { trace.push("first"); return "1"; })()} b=${boom()} c=${"never"}`);
} catch (e) {
    console.log(e.message);
    console.log(trace.join(","));
}
