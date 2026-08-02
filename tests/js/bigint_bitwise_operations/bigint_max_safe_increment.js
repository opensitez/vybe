// vybe-test: js/bigint_bitwise_operations/bigint_max_safe_increment
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

const b=9007199254740991n; __check(__line((b+1n).toString()), "9007199254740992");
