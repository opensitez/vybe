// vybe-test: js/ecma_objects/object_literal_numeric_keys_access_as_strings
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const obj = { 1: "one", 2: "two" };
__check(__line(obj[1]), "one");
__check(__line(obj["2"]), "two");
