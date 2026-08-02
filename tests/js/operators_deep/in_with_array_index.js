// vybe-test: js/operators_deep/in_with_array_index
// origin: languages/js/tests/js/test_operators_deep.rs

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

const arr = [1, 2, 3];
__check(__line(0 in arr), "true");
__check(__line(2 in arr), "true");
__check(__line(3 in arr), "false");  // out of bounds
