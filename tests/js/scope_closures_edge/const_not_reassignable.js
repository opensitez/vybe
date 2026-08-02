// vybe-test: js/scope_closures_edge/const_not_reassignable
// origin: languages/js/tests/js/test_scope_closures_edge.rs

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

const x = 42;
let threw = false;
try { eval("const x = 42; x = 1;"); } catch { threw = true; }
__check(__line(threw), "true");
