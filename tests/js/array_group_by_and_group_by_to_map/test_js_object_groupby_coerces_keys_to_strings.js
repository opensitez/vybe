// vybe-test: js/array_group_by_and_group_by_to_map/test_js_object_groupby_coerces_keys_to_strings
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

const numbers = [10, 20, 30];
const grouped = Object.groupBy(numbers, x => x);
__check(__line(Object.keys(grouped).join(",")), "10,20,30");
