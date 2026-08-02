// vybe-test: js/map_set_advanced_patterns/object_groupby_pattern
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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

const people = [
    { name: "Alice", dept: "eng" },
    { name: "Bob", dept: "hr" },
    { name: "Charlie", dept: "eng" },
];
const grouped = Object.groupBy(people, p => p.dept);
__check(__line(grouped.eng.length), "2");
__check(__line(grouped.hr.length), "1");
__check(__line(grouped.eng[0].name), "Alice");
