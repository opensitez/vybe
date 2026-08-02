// vybe-test: js/comprehensive/test_in_operator_computed
// origin: languages/js/tests/js/js_comprehensive_test.rs

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

let obj = { x: 1, y: 2 };
        let key = "x";
        __check(__line(key in obj), "true");
