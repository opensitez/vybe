// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_declaration_in_if_block_annex_b_non_strict
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

if (true) {
    function ifFunc() { return "IfBlockFunc"; }
}
__check(__line(ifFunc()), "IfBlockFunc"); // In non-strict Annex B, function declaration in block hoists to enclosing function/global!
