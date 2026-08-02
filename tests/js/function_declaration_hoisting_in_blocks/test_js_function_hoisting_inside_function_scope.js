// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_hoisting_inside_function_scope
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

function outer() {
    return inner();
    function inner() { return "InnerHoisted"; }
}
__check(__line(outer()), "InnerHoisted");
