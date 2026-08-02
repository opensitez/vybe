// vybe-test: js/array_group_by_and_group_by_to_map/test_js_map_groupby_object_keys
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

const rest1 = { name: "RestA" };
const rest2 = { name: "RestB" };
const goods = [
    { item: "apple", rest: rest1 },
    { item: "pear", rest: rest1 },
    { item: "steak", rest: rest2 }
];
const groupedMap = Map.groupBy(goods, g => g.rest);
__check(__line((groupedMap instanceof Map) + "|" + groupedMap.get(rest1).length + "|" + groupedMap.get(rest2).length), "true|2|1");
