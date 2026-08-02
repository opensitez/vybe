// vybe-test: js/ecma/test_destructure_object_rename
// origin: languages/js/tests/js/js_ecma_test.rs

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

let obj = { name: "Alice", age: 30 };
        let { name: n, age: a } = obj;
        __check(__line(n, a), "Alice 30");
