// vybe-test: js/map_groupby_object_groupby_utilities/test_js_object_groupby_set_iterable
// origin: languages/js/tests/js/test_js_map_groupby_object_groupby_utilities.rs

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

const set = new Set([1, 2, 3, 4]);
const grouped = Object.groupBy(set, x => x > 2 ? "high" : "low");
__check(__line(grouped.low.join(",") + "|" + grouped.high.join(",")), "1,2|3,4");
