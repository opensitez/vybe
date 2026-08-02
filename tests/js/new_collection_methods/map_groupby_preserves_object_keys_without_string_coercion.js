// vybe-test: js/new_collection_methods/map_groupby_preserves_object_keys_without_string_coercion
// origin: languages/js/tests/js/test_new_collection_methods.rs

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

const low = { label: "low" };
const high = { label: "high" };
const grouped = Map.groupBy([1, 2, 3], value => value < 3 ? low : high);
__check(__line(grouped.get(low).join(",")), "1,2");
__check(__line(grouped.get(high).join(",")), "3");
__check(__line(grouped.has({ label: "low" })), "false");
