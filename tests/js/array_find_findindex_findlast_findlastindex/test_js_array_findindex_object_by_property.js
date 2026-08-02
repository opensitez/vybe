// vybe-test: js/array_find_findindex_findlast_findlastindex/test_js_array_findindex_object_by_property
// origin: languages/js/tests/js/test_js_array_find_findindex_findlast_findlastindex.rs

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

const inventory = [
    { name: "apples", quantity: 2 },
    { name: "bananas", quantity: 0 }
];
const idx = inventory.findIndex(item => item.name === "bananas");
__check(__line(idx), "1");
