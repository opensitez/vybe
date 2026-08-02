// vybe-test: js/es2023_2025_features/map_group_by
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

const words = ["apple", "banana", "avocado", "blueberry"];
const grouped = Map.groupBy(words, w => w[0]);
__check(__line(grouped.get("a").join(",")), "apple,avocado");
__check(__line(grouped.get("b").join(",")), "banana,blueberry");
