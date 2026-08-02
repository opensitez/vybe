// vybe-test: js/closure_scope_deep_patterns/scope_resolution_lexical
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

const x = "global";
function outer() {
    const x = "outer";
    function inner() {
        return x;  // lexically captured
    }
    return inner;
}
const fn = outer();
__check(__line(fn()), "outer");  // "outer" not "global"
