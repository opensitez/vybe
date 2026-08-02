// vybe-test: js/array_destructuring_elision_rest_element/test_js_array_destructuring_rest_element_must_be_last
// origin: languages/js/tests/js/test_js_array_destructuring_elision_rest_element.rs

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
    eval("const [...rest, last] = [1, 2];");
} catch (e) {
    __check(__line("Rest Element Not Last SyntaxError"), "Rest Element Not Last SyntaxError");
}
