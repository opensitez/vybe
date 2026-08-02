// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_nested_block_shadowing
// origin: languages/js/tests/js/test_js_block_scoping_let_const_in_loops_and_blocks.rs

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
    const x = "middle";
    {
        const x = "inner";
        __check(__line(x), "inner");
    }
    __check(__line(x), "middle");
}
__check(__line(x), "outer");
