// vybe-test: js/operators_deep/left_shift_right_shift
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

__check(__line(1 << 4), "16");   // 16
__check(__line(16 >> 2), "4");  // 4
__check(__line(-1 >> 1), "-1");  // -1 (sign preserved)
__check(__line(-1 >>> 1), "2147483647"); // 2147483647 (unsigned)
