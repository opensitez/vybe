// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_declaration_in_while_loop_block_annex_b
// origin: languages/js/tests/js/test_js_function_declaration_hoisting_in_blocks.rs

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
while (i < 1) {
    function whileFunc() { return "WhileFunc"; }
    i++;
}
console.log(whileFunc());
