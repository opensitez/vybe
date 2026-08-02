// vybe-test: js/stdlib/test_arr_slice
// origin: languages/js/tests/js/js_stdlib_test.rs

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

let a = [1, 2, 3, 4, 5]; __check(__line(a.slice(1, 3)), "2,3")
