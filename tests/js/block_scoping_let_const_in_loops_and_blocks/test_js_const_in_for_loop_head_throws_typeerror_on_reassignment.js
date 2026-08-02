// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_const_in_for_loop_head_throws_typeerror_on_reassignment
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

try {
    eval("for (const i = 0; i < 2; i++) {}"); // i++ attempts to reassign const in loop head!
} catch (e) {
    __check(__line("For Loop Const TypeError"), "For Loop Const TypeError");
}
