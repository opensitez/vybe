// vybe-test: js/new_collection_methods/map_groupby_groups_elements_by_string_key
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

const entries = Map.groupBy([1, 2, 3, 4], value => value % 2 === 0 ? "even" : "odd");
__check(__line(entries.get("even").join(",")), "2,4");
__check(__line(entries.get("odd").join(",")), "1,3");
