// vybe-test: js/set_methods_union_intersection_difference/test_js_set_methods_missing_has_method_throws
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

const s = new Set([1]);
const invalidSetLike = { size: 1, keys() { return [1][Symbol.iterator](); } };
try {
    s.intersection(invalidSetLike);
} catch (e) {
    __check(__line("Set Method Invalid Set-Like TypeError"), "Set Method Invalid Set-Like TypeError");
}
