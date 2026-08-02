// vybe-test: js/arrow_functions/arrow_expression_body_undefined_vs_empty_block
// origin: languages/js/tests/js/test_arrow_functions.rs

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

const f1 = () => undefined; const f2 = () => {}; __check(__line(f1() === undefined && f2() === undefined), "true");
