// vybe-test: js/number_bigint/bigint_as_uint_n_clamps_to_unsigned_n_bit
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

// 256 mod 2^8 = 0
__check(__line(BigInt.asUintN(8, 256n)), "0n");
// 255 fits exactly in 8-bit unsigned
__check(__line(BigInt.asUintN(8, 255n)), "255n");
