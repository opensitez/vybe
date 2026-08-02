// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flat_zero_depth_returns_shallow_copy
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

const arr = [1, [2]];
const f0 = arr.flat(0);
__check(__line(f0.length + "|" + (f0 !== arr)), "2|true");
