// vybe-test: js/ecma_functions/arrow_returns_object
// origin: languages/js/tests/js/test_ecma_functions.rs

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

const makeObj = (k, v) => ({ [k]: v });
const obj = makeObj("name", "Alice");
__check(__line(obj.name), "Alice");
