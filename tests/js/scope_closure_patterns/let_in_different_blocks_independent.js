// vybe-test: js/scope_closure_patterns/let_in_different_blocks_independent
// origin: languages/js/tests/js/test_scope_closure_patterns.rs

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

let x = "outer";
{
    let x = "inner1";
    __check(__line(x), "inner1");
}
{
    let x = "inner2";
    __check(__line(x), "inner2");
}
__check(__line(x), "outer");
