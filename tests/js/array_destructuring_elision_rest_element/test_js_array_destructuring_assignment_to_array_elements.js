// vybe-test: js/array_destructuring_elision_rest_element/test_js_array_destructuring_assignment_to_array_elements
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

const arr = [0, 0];
[arr[0], arr[1]] = [5, 15];
__check(__line(arr.join(",")), "5,15");
