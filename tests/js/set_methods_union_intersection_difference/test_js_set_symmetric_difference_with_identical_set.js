// vybe-test: js/set_methods_union_intersection_difference/test_js_set_symmetric_difference_with_identical_set
// origin: languages/js/tests/js/test_js_set_methods_union_intersection_difference.rs

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

const s1 = new Set(["a", "b"]);
const s2 = new Set(["a", "b"]);
__check(__line(s1.symmetricDifference(s2).size), "0");
