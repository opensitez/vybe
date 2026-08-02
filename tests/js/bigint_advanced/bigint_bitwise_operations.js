// vybe-test: js/bigint_advanced/bigint_bitwise_operations
// origin: languages/js/tests/js/test_bigint_advanced.rs

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

__check(__line((0b1100n & 0b1010n).toString()), "8");
__check(__line((0b1100n | 0b1010n).toString()), "14");
__check(__line((0b1100n ^ 0b1010n).toString()), "6");
__check(__line((~0b1100n).toString()), "-13"); // two's complement
