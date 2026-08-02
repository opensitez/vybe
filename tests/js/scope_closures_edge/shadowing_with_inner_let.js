// vybe-test: js/scope_closures_edge/shadowing_with_inner_let
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

const x = "outer";
{
    const x = "inner";
    __check(__line(x), "inner");
}
__check(__line(x), "outer");
