// vybe-test: js/es2023_2025_features/array_group_by_object
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

const items = [
    { type: "A", val: 1 }, { type: "B", val: 2 }, { type: "A", val: 3 }
];
const grouped = Object.groupBy(items, item => item.type);
__check(__line(grouped.A.length), "2");
__check(__line(grouped.B.length), "1");
__check(__line(grouped.A[0].val), "1");
