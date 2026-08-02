// vybe-test: js/set_operations_es2025/empty_set_is_subset_of_any_set
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

const empty = new Set();
const a = new Set([1, 2, 3]);
__check(__line(empty.isSubsetOf(a)), "true");
const empty2 = new Set();
__check(__line(empty2.isSubsetOf(new Set([10, 20]))), "true");
