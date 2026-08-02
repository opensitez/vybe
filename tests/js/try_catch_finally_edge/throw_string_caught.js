// vybe-test: js/try_catch_finally_edge/throw_string_caught
// origin: languages/js/tests/js/test_try_catch_finally_edge.rs

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

try {
    throw "string error";
} catch (e) {
    __check(__line(typeof e), "string");
    __check(__line(e), "string error");
}
