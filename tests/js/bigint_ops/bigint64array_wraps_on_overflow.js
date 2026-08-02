// vybe-test: js/bigint_ops/bigint64array_wraps_on_overflow
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

const arr = new BigInt64Array(1);
const max = 9223372036854775807n;
arr[0] = max + 1n;
__check(__line(arr[0]), "-9223372036854775808n");
