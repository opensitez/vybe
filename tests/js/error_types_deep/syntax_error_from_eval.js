// vybe-test: js/error_types_deep/syntax_error_from_eval
// origin: languages/js/tests/js/test_error_types_deep.rs

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

let err = null;
try { eval("if ("); } catch (e) { err = e; }
__check(__line(err instanceof SyntaxError), "true");
__check(__line(err instanceof Error), "true");
