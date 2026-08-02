// vybe-test: js/scope_closures_edge/let_tdz_throws_before_declaration
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

let threw = false;
try {
    eval("console.log(y); let y = 1;");
} catch (e) {
    threw = e instanceof ReferenceError;
}
__check(__line(threw), "true");
