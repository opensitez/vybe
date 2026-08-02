// vybe-test: js/comprehensive/test_arrow_block_body
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

let f = (a, b) => {
            let sum = a + b;
            return sum;
        };
        __check(__line(f(3, 7)), "10");
