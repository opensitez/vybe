// vybe-test: js/comprehensive/test_destructuring_nested_object
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

let { a, b: { c } } = { a: 1, b: { c: 2 } };
        __check(__line(a, c), "1 2");
