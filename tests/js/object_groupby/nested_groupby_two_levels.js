// vybe-test: js/object_groupby/nested_groupby_two_levels
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

const data = [
    { dept: "eng", level: "senior" },
    { dept: "eng", level: "junior" },
    { dept: "hr", level: "senior" }
];
const byDept = Object.groupBy(data, d => d.dept);
const engByLevel = Object.groupBy(byDept.eng, d => d.level);
__check(__line(engByLevel.senior.length), "1");
__check(__line(engByLevel.junior.length), "1");
