// vybe-test: js/number_bigint/bigint_as_int_n_clamps_to_signed_n_bit
// origin: languages/js/tests/js/test_number_bigint.rs

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

// 128 in signed 8-bit wraps to -128
__check(__line(BigInt.asIntN(8, 128n)), "-128n");
// 127 fits in signed 8-bit
__check(__line(BigInt.asIntN(8, 127n)), "127n");
