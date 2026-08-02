// vybe-test: js/interop/test_b16_spread_in_object
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

let other = { a: 1, b: 2 };
        let obj = { ...other, c: 3 };
        __check(__line(obj.a, obj.b, obj.c), "1 2 3");
