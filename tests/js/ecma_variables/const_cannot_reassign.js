// vybe-test: js/ecma_variables/const_cannot_reassign
// origin: languages/js/tests/js/test_ecma_variables.rs

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

const x = 42;
__check(__line(x), "42");
