// vybe-test: js/map_groupby_object_groupby_utilities/test_js_object_groupby_numeric_or_string_keys
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

const numbers = [1, 2, 3, 4, 5, 6];
const result = Object.groupBy(numbers, num => num % 2 === 0 ? "even" : "odd");
__check(__line(result.even.join(",") + "|" + result.odd.join(",")), "2,4,6|1,3,5");
