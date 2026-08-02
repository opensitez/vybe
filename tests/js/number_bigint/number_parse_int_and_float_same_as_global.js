// vybe-test: js/number_bigint/number_parse_int_and_float_same_as_global
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

__check(__line(Number.parseInt("10") === parseInt("10")), "true");
__check(__line(Number.parseFloat("3.14") === parseFloat("3.14")), "true");
