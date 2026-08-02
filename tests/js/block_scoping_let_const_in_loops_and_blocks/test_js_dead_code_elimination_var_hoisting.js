// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_dead_code_elimination_var_hoisting
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

if (false) {
    var unreachedVar = 1;
}
__check(__line(unreachedVar), "undefined"); // var declaration is hoisted even inside unreached if (false) block!
