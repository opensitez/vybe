// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_string_keys
// origin: languages/js/tests/js/test_js_array_group_by_and_group_by_to_map.rs

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
    { name: "asparagus", type: "vegetable" },
    { name: "bananas", type: "fruit" },
    { name: "goat", type: "meat" },
    { name: "cherries", type: "fruit" }
];
const result = Object.groupBy(inventory, item => item.type);
__check(__line(`${result.vegetable.length}:${result.fruit.length}:${result.meat.length}`), "1:2:1");
