// vybe-test: js/bigint_ops/bigint64array_stores_bigint_values
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

const arr = new BigInt64Array(3);
arr[0] = 100n;
arr[1] = -200n;
arr[2] = 9007199254740993n;
__check(__line(arr[0]), "100n");
__check(__line(arr[1]), "-200n");
__check(__line(arr[2]), "9007199254740993n");
