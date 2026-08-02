// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flat_non_array_elements_not_flattened
// origin: languages/js/tests/js/test_js_array_flat_flatmap_depth_recursion.rs

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

const arrayLikeObject = { 0: "a", 1: "b", length: 2 };
const arr = [1, arrayLikeObject];
const flattened = arr.flat();
__check(__line(flattened.length + "|" + (typeof flattened[1])), "2|object");
