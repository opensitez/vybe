// vybe-test: js/set_methods_union_intersection_difference/test_js_set_methods_returns_new_set_instance
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
const u = s1.union(s2);
__check(__line((u instanceof Set) + "|" + (u !== s1) + "|" + (u !== s2)), "true|true|true");
