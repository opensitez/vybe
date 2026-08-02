// vybe-test: js/misc_es_features/nullish_assignment_operator
// origin: languages/js/tests/js/test_misc_es_features.rs

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

let a = null;
a ??= "assigned";
__check(__line(a), "assigned");
let b = "existing";
b ??= "not assigned";
__check(__line(b), "existing");
