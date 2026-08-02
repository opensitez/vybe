// vybe-test: js/array_flat_flatmap_depth_recursion/test_js_array_flatmap_returning_non_array_iterable
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

const set = new Set([10, 20]);
const res = [1].flatMap(() => set);
__check(__line(res.join(",") + "|isArr=" + Array.isArray(res)), "10,20|isArr=true"); // Returns flattened Array, not Set!
