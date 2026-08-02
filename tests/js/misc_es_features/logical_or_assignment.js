// vybe-test: js/misc_es_features/logical_or_assignment
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

let a = 0;
a ||= 5;
__check(__line(a), "5");
let b = 3;
b ||= 5;
__check(__line(b), "3");
