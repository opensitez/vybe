// vybe-test: js/bigint_ops/number_to_bigint_explicit_conversion
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

const n = BigInt(Number.MAX_SAFE_INTEGER);
__check(__line(n === 9007199254740991n), "true");
