// vybe-test: js/error_hierarchy/error_wrapping_pattern
// origin: languages/js/tests/js/test_error_hierarchy.rs

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

function parse(str) {
    try {
        return JSON.parse(str);
    } catch (e) {
        throw new Error("Parse failed: " + e.message, { cause: e });
    }
}

try {
    parse("{bad}");
} catch (e) {
    __check(__line(e.message.startsWith("Parse failed:")), "true");
    __check(__line(e.cause instanceof SyntaxError), "true");
}
