// vybe-test: js/nested_template_literals_evaluation/test_js_nested_template_literal_try_catch_block
// origin: languages/js/tests/js/test_js_nested_template_literals_evaluation.rs

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

let res;
try {
    res = `Try_${`Success_${10}`}`;
} catch (e) {
    res = `Catch_${`Error`}`;
}
__check(__line(res), "Try_Success_10");
