// vybe-test: js/array_destructuring_elision_rest_element/test_js_array_destructuring_assignment_to_object_properties
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

const obj = {};
[obj.x, obj.y] = [10, 20];
__check(__line(`${obj.x}:${obj.y}`), "10:20");
