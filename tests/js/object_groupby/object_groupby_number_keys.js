// vybe-test: js/object_groupby/object_groupby_number_keys
// origin: languages/js/tests/js/test_object_groupby.rs

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

const nums = [10, 21, 35, 42, 57];
const groups = Object.groupBy(nums, n => Math.floor(n / 10) * 10);
const keys = Object.keys(groups).sort((a, b) => +a - +b);
__check(__line(keys.join(",")), "10,20,30,40,50");
