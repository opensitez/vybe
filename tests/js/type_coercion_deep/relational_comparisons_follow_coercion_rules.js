// vybe-test: js/type_coercion_deep/relational_comparisons_follow_coercion_rules
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

__check(__line([1] > 0), "true");       // array -> string -> number
__check(__line([1, 2] > 0), "false");    // array with >1 element coerces to NaN
__check(__line(true > 0), "true");
__check(__line(false < 1), "true");
