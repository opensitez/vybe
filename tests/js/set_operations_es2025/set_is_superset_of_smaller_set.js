// vybe-test: js/set_operations_es2025/set_is_superset_of_smaller_set
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

const a = new Set([1, 2, 3, 4]);
const b = new Set([2, 3]);
__check(__line(a.isSupersetOf(b)), "true");
__check(__line(b.isSupersetOf(a)), "false");
