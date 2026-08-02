// vybe-test: js/new_collection_methods/map_groupby_tracks_number_of_groups
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

const grouped = Map.groupBy(["ant", "bear", "cat"], value => value.length);
__check(__line(grouped.size), "2");
__check(__line(grouped.get(3).join(",")), "ant,cat");
__check(__line(grouped.get(4).join(",")), "bear");
