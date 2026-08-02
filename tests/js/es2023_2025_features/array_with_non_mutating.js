// vybe-test: js/es2023_2025_features/array_with_non_mutating
// origin: languages/js/tests/js/test_es2023_2025_features.rs

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

const orig = [1, 2, 3];
const updated = orig.with(1, 99);
__check(__line(updated.join(",")), "1,99,3");
__check(__line(orig.join(",")), "1,2,3");
