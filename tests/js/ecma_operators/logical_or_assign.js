// vybe-test: js/ecma_operators/logical_or_assign
// origin: languages/js/tests/js/test_ecma_operators.rs

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
let b = 1;
a ||= 42;
b ||= 42;
__check(__line(a), "42");
__check(__line(b), "1");
