// vybe-test: js/map_groupby_object_groupby_utilities/test_js_map_groupby_complex_object_keys
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

const keyEven = { type: "even" };
const keyOdd = { type: "odd" };
const numbers = [10, 15, 20, 25];

const map = Map.groupBy(numbers, n => n % 2 === 0 ? keyEven : keyOdd);
__check(__line(map.get(keyEven).join(",") + "|" + map.get(keyOdd).join(",")), "10,20|15,25");
