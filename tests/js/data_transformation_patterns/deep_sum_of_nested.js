// vybe-test: js/data_transformation_patterns/deep_sum_of_nested
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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

function deepSum(obj) {
    if (typeof obj === "number") return obj;
    if (Array.isArray(obj)) return obj.reduce((s, x) => s + deepSum(x), 0);
    if (typeof obj === "object") return Object.values(obj).reduce((s, v) => s + deepSum(v), 0);
    return 0;
}
const data = { a: 1, b: [2, 3, { c: 4 }], d: { e: 5, f: [6, 7] } };
__check(__line(deepSum(data)), "28");
