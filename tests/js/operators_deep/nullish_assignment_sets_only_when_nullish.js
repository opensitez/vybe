// vybe-test: js/operators_deep/nullish_assignment_sets_only_when_nullish
// origin: languages/js/tests/js/test_operators_deep.rs

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

let a;
let b = 0;
a ??= 5;
b ??= 7;
__check(__line(a), "5");
__check(__line(b), "0");
