// vybe-test: js/misc_es_features/logical_and_assignment
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

let a = 1;
a &&= 2;
__check(__line(a), "2");
let b = 0;
b &&= 2;
__check(__line(b), "0");
