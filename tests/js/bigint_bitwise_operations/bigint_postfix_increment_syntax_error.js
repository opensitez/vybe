// vybe-test: js/bigint_bitwise_operations/bigint_postfix_increment_syntax_error
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

let b=1n; b++; __check(__line(b), "2n");
