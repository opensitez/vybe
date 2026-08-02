// vybe-test: js/scope_closure_patterns/function_expression_in_block
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

{
    const fn = function named() { return 42; };
    __check(__line(fn()), "42");
    __check(__line(fn.name), "named");
}
