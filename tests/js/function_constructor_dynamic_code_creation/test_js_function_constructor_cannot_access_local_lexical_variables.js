// vybe-test: js/function_constructor_dynamic_code_creation/test_js_function_constructor_cannot_access_local_lexical_variables
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

(() => {
    const hiddenVal = 99;
    const fn = new Function("try { return hiddenVal; } catch(e) { return 'ReferenceError'; }");
    __check(__line(fn()), "ReferenceError");
})();
