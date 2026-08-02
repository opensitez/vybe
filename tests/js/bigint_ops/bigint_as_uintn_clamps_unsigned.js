// vybe-test: js/bigint_ops/bigint_as_uintn_clamps_unsigned
// origin: languages/js/tests/js/test_bigint_ops.rs

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

__check(__line(BigInt.asUintN(8, 256n)), "0n");
__check(__line(BigInt.asUintN(8, 255n)), "255n");
