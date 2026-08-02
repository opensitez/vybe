// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_let_tdz_in_same_scope_before_initialization_throws
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

let hit = "no";
try {
    {
        console.log(x);
        let x = "too late";
    }
} catch (e) {
    hit = "error";
}
console.log(hit);
