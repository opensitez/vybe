// vybe-test: js/bigint_bitwise_operations/bigint_abs_via_condition
// origin: languages/js/tests/js/test_bigint_bitwise_operations.rs

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

const b=-9n; __check(__line(b<0n?-b:b), "9n");
