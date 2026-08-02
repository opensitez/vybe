// vybe-test: js/map_groupby_object_groupby_utilities/test_js_map_groupby_preserve_insertion_order_of_keys
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

const items = [{ cat: "A", id: 1 }, { cat: "B", id: 2 }, { cat: "A", id: 3 }];
const grouped = Map.groupBy(items, i => i.cat);
__check(__line([...grouped.keys()].join(",")), "A,B");
