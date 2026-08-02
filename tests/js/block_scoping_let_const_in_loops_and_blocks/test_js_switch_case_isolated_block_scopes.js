// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_switch_case_isolated_block_scopes
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

switch(1) {
    case 1: {
        let a = 1;
        console.log(a);
        break;
    }
    case 2: {
        let a = 2;
        console.log(a);
        break;
    }
}
