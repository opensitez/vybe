// vybe-test: js/object_groupby/groupby_range_buckets
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

const scores = [45, 67, 89, 92, 78, 55, 33];
const grades = Object.groupBy(scores, s => {
    if (s >= 90) return "A";
    if (s >= 70) return "B";
    if (s >= 50) return "C";
    return "F";
});
__check(__line(grades.A.join(",")), "92");
__check(__line(grades.B.join(",")), "89,78");
__check(__line(grades.F.join(",")), "45,33");
