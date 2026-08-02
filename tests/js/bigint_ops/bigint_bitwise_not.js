// vybe-test: js/bigint_ops/bigint_bitwise_not
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

__check(__line(~0n), "-1n");
__check(__line(~1n), "-2n");
__check(__line(~(-1n)), "0n");
