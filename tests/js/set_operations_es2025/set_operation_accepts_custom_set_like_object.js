// vybe-test: js/set_operations_es2025/set_operation_accepts_custom_set_like_object
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

const a = new Set([1, 2, 3]);
const custom = {
    size: 2,
    has(v) { return v === 2 || v === 3; },
    keys() { return [2, 3][Symbol.iterator](); }
};
__check(__line([...a.intersection(custom)].join(",")), "2,3");
