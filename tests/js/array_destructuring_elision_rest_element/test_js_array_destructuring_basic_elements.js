// vybe-test: js/array_destructuring_elision_rest_element/test_js_array_destructuring_basic_elements
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

const [a, b, c] = [10, 20, 30];
__check(__line(`${a},${b},${c}`), "10,20,30");
