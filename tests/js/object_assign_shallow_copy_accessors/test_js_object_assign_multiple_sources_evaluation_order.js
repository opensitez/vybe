// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_multiple_sources_evaluation_order
// origin: languages/js/tests/js/test_js_object_assign_shallow_copy_accessors.rs

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

const log = [];
const s1 = { get a() { log.push("s1.a"); return 1; } };
const s2 = { get a() { log.push("s2.a"); return 2; } };
const target = Object.assign({}, s1, s2);
__check(__line(target.a + "|" + log.join(",")), "2|s1.a,s2.a");
