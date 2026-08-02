// vybe-test: js/function_declaration_hoisting_in_blocks/test_js_function_declaration_with_default_parameters_evaluation_order
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

function fn(a = getA()) {
    function getA() { return 100; } // getA is hoisted within body scope, NOT accessible to parameter default evaluation!
    return a;
}
try {
    fn();
} catch (e) {
    __check(__line("Parameter Default Function Hoisting ReferenceError"), "Parameter Default Function Hoisting ReferenceError");
}
