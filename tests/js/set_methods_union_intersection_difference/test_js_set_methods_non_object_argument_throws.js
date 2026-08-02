// vybe-test: js/set_methods_union_intersection_difference/test_js_set_methods_non_object_argument_throws
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
try {
    s.union(12345);
} catch (e) {
    __check(__line("Set Method Non-Object TypeError"), "Set Method Non-Object TypeError");
}
