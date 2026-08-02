// vybe-test: js/number_bigint/bigint_mixed_arithmetic_with_number_throws
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

try {
    let result = 1n + 1;
    console.log("no error");
} catch (e) {
    console.log(e instanceof TypeError);
}
