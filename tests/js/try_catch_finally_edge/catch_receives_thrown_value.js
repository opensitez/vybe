// vybe-test: js/try_catch_finally_edge/catch_receives_thrown_value
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
    throw { code: 42, msg: "custom" };
} catch (e) {
    __check(__line(e.code), "42");
    __check(__line(e.msg), "custom");
}
