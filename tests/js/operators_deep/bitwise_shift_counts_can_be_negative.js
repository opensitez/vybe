// vybe-test: js/operators_deep/bitwise_shift_counts_can_be_negative
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

__check(__line(1 << -1), "-2147483648");   // -1 -> 31
__check(__line(-1 << -1), "-2147483648");  // -1 -> 31
__check(__line(1 >> -1), "0");   // same as >>31
__check(__line(-1 >> -1), "-1");
