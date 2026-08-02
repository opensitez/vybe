// vybe-test: js/es2023_2025_features/object_has_own
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

const obj = { a: 1, b: 2 };
__check(__line(Object.hasOwn(obj, "a")), "true");
__check(__line(Object.hasOwn(obj, "toString")), "false");
const n = Object.create(null);
n.x = 5;
__check(__line(Object.hasOwn(n, "x")), "true");
