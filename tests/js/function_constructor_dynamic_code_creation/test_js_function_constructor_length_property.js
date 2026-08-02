// vybe-test: js/function_constructor_dynamic_code_creation/test_js_function_constructor_length_property
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

const fn1 = new Function("a", "b", "c", "return 0;");
const fn2 = new Function("a, b = 1", "c", "return 0;");
__check(__line(`${fn1.length}:${fn2.length}`), "3:1");
