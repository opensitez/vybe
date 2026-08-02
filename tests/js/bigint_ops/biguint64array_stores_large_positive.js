// vybe-test: js/bigint_ops/biguint64array_stores_large_positive
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

const arr = new BigUint64Array(1);
arr[0] = 18446744073709551615n;
__check(__line(arr[0]), "18446744073709551615n");
