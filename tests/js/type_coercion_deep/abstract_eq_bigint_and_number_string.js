// vybe-test: js/type_coercion_deep/abstract_eq_bigint_and_number_string
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

__check(__line(1n == 1), "true");
__check(__line(1n == "1"), "true");
__check(__line(0n == false), "true");
__check(__line(2n == 2.5), "false");
__check(__line(1n === 1), "false");
