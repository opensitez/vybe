// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_redefinition_throws_syntaxerror
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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

try {
    eval("class Dup { #fn() {} #fn() {} }");
} catch (e) {
    __check(__line("Duplicate Private Method SyntaxError"), "Duplicate Private Method SyntaxError");
}
