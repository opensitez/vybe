// vybe-test: js/array_higher_order/array_map_groupby_primitive_keys
// origin: languages/js/tests/js/test_array_higher_order.rs

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

const nums = [1, 2, 3, 4, 5];
const grouped = Map.groupBy(nums, x => x % 2 === 0 ? "even" : "odd");
__check(__line(grouped.get("even").join(",") + "|" + grouped.get("odd").join(",")), "2,4|1,3,5");
