// vybe-test: js/ecma/test_spread_in_array
// origin: languages/js/tests/js/js_ecma_test.rs

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

let a = [1, 2];
        let b = [3, 4];
        let c = [...a, ...b, 5];
        __check(__line(c.join(",")), "1,2,3,4,5");
