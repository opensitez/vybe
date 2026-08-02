// vybe-test: js/set_methods_union_intersection_difference/test_js_set_union_chaining
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

const s1 = new Set([1]);
const s2 = new Set([2]);
const s3 = new Set([3]);
const all = s1.union(s2).union(s3);
__check(__line([...all].join(",")), "1,2,3");
