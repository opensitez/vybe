// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_named_function_expression_name_only_in_own_scope
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

const fn = function internalName() {
    return typeof internalName;
};
__check(__line(fn() + "|" + (typeof internalName)), "function|undefined");
