// vybe-test: js/map_set_advanced_patterns/map_groupby_pattern
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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

const items = [1, 2, 3, 4, 5, 6];
const grouped = Map.groupBy(items, x => x % 2 === 0 ? "even" : "odd");
__check(__line(grouped.get("even").join(",")), "2,4,6");
__check(__line(grouped.get("odd").join(",")), "1,3,5");
