// vybe-test: js/function_constructor_dynamic_code_creation/test_js_function_constructor_created_function_can_be_used_as_constructor
// origin: languages/js/tests/js/test_js_function_constructor_dynamic_code_creation.rs

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

const MyClass = new Function("val", "this.val = val;");
const inst = new MyClass("DynamicInst");
__check(__line(inst.val), "DynamicInst");
