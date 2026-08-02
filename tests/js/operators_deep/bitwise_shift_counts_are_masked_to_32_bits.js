// vybe-test: js/operators_deep/bitwise_shift_counts_are_masked_to_32_bits
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

__check(__line(1 << 32), "1");
__check(__line(1 << 33), "2");
__check(__line(1 << 40), "256");
__check(__line(-1 << 1), "-2");
