// vybe-test: js/object_groupby/groupby_then_map_groups
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

const data = [1, 2, 3, 4, 5, 6];
const groups = Object.groupBy(data, n => n % 2 === 0 ? "even" : "odd");
const sums = {};
for (const [key, vals] of Object.entries(groups)) {
    sums[key] = vals.reduce((a, b) => a + b, 0);
}
console.log(sums.even);
console.log(sums.odd);
