// vybe-test: js/array_methods_new/with_out_of_bounds_index_throws_rangeerror
// origin: languages/js/tests/js/test_array_methods_new.rs

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

try {
    [1, 2, 3].with(10, 99);
    console.log("no error");
} catch (e) {
    console.log(e instanceof RangeError);
}
