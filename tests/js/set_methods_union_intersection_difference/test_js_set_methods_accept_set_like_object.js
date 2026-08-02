// vybe-test: js/set_methods_union_intersection_difference/test_js_set_methods_accept_set_like_object
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

const s1 = new Set([1, 2, 3]);
const setLike = {
    size: 2,
    has(v) { return v === 2 || v === 3; },
    keys() { return [2, 3][Symbol.iterator](); }
};
const i = s1.intersection(setLike);
__check(__line([...i].join(",")), "2,3");
