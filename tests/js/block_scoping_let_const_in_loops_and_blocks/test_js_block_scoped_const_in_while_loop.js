// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_block_scoped_const_in_while_loop
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

let i = 0;
while (i < 2) {
    const loopConst = i * 10;
    console.log(loopConst);
    i++;
}
