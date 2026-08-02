// vybe-test: js/ecma_variables/destructure_object_rename
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

const obj = { name: "Alice", age: 30 };
const { name: n, age: a } = obj;
__check(__line(n), "Alice");
__check(__line(a), "30");
