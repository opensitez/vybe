// vybe-test: js/interop/test_c32_array_from_creates_copy
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

let a = [1, 2, 3];
        let b = Array.from(a);
        b[0] = 99;
        __check(__line(a[0], b[0]), "1 99");
