// vybe-test: js/map_groupby_object_groupby_utilities/test_js_map_groupby_supports_symbol_keys
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

const symKey = Symbol("sym");
const grouped = Map.groupBy(["a", "b"], () => symKey);
__check(__line(grouped.get(symKey).join(",")), "a,b");
