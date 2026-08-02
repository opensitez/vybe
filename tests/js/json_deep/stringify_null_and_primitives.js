// vybe-test: js/json_deep/stringify_null_and_primitives
// origin: languages/js/tests/js/test_json_deep.rs

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

__check(__line(JSON.stringify(null)), "null");
__check(__line(JSON.stringify(42)), "42");
__check(__line(JSON.stringify("hello")), "\"hello\"");
__check(__line(JSON.stringify(true)), "true");
