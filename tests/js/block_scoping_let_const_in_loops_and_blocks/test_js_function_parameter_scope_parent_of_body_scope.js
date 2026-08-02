// vybe-test: js/block_scoping_let_const_in_loops_and_blocks/test_js_function_parameter_scope_parent_of_body_scope
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

function fn(param = "defaultParam") {
    let paramVar = "bodyVar";
    __check(__line(`${param}:${paramVar}`), "defaultParam:bodyVar");
}
fn();
