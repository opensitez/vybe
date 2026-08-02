// vybe-test: js/interop/test_g70_template_literal_with_expression
// origin: languages/js/tests/js/js_interop_test.rs

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

let a = 3;
        let b = 4;
        __check(__line(`${a} + ${b} = ${a + b}`), "3 + 4 = 7");
