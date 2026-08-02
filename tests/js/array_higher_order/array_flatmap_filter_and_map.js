// vybe-test: js/array_higher_order/array_flatmap_filter_and_map
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

// flatMap can filter by returning [] for excluded items
const nums = [1, 2, 3, 4, 5];
const evenDoubled = nums.flatMap(x => x % 2 === 0 ? [x * 2] : []);
console.log(evenDoubled.join(","));
