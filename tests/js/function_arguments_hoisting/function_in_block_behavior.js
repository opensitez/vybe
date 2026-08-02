// vybe-test: js/function_arguments_hoisting/function_in_block_behavior
// origin: languages/js/tests/js/test_function_arguments_hoisting.rs

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

// Block-scoped function declaration (non-strict behavior: hoisted as var)
__check(__line(typeof blockFn), "undefined");
{
    function blockFn() { return "inside"; }
}
__check(__line(typeof blockFn), "function");
