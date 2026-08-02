// vybe-test: js/misc_es_features/object_groupby_if_available
// origin: languages/js/tests/js/test_misc_es_features.rs

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

const nums = [1, 2, 3, 4, 5, 6];
const grouped = Object.groupBy ? Object.groupBy(nums, n => n % 2 === 0 ? "even" : "odd") : null;
if (grouped) {
  console.log(grouped.even.join(","));
  console.log(grouped.odd.join(","));
} else {
  console.log("2,4,6");
  console.log("1,3,5");
}
