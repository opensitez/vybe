// vybe-test: js/scope_tdz/catch_binding_scoped_to_catch_block_var_escapes
// origin: languages/js/tests/js/test_scope_tdz.rs

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

try { throw new Error("test"); }
catch (e) { var caught = e.message; }
__check(__line(caught), "test");
__check(__line(typeof e), "undefined");
