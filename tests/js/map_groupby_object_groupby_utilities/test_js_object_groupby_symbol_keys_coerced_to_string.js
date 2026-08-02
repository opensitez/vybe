// vybe-test: js/map_groupby_object_groupby_utilities/test_js_object_groupby_symbol_keys_coerced_to_string
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

const items = ["apple", "apricot", "banana"];
const grouped = Object.groupBy(items, item => item[0]);
__check(__line(grouped["a"].join(",") + "|" + grouped["b"].join(",")), "apple,apricot|banana");
