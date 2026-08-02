// vybe-test: js/number_bigint/number_epsilon_is_positive_and_less_than_one
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

__check(__line(Number.EPSILON > 0), "true");
__check(__line(Number.EPSILON < 1), "true");
