// vybe-test: js/error_types/error_chaining
// origin: languages/js/tests/js/test_error_types.rs

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

function inner() {
    throw new Error("from inner");
}
function middle() {
    try {
        inner();
    } catch (e) {
        throw new Error("from middle: " + e.message);
    }
}
try {
    middle();
} catch (e) {
    __check(__line(e.message), "from middle: from inner");
}
