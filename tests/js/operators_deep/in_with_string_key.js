// vybe-test: js/operators_deep/in_with_string_key
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

const obj = { a: undefined }; // key exists but value undefined
__check(__line("a" in obj), "true");
__check(__line("b" in obj), "false");
__check(__line(obj.a === undefined), "true"); // exists but undefined
