// vybe-test: js/set_operations_es2025/empty_sets_are_disjoint
// origin: languages/js/tests/js/test_set_operations_es2025.rs

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

const a = new Set();
const b = new Set([1, 2]);
__check(__line(a.isDisjointFrom(b)), "true");
__check(__line(a.isDisjointFrom(new Set())), "true");
