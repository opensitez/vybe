// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flat_custom_depth
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

const nested = [1, [2, [3, [4]]]];
const f2 = nested.flat(2);
__check(__line(f2.map(x => Array.isArray(x) ? x.join(":") : x).join(",")), "1,2,3,4");
