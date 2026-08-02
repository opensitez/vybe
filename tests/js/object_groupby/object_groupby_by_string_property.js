// vybe-test: js/object_groupby/object_groupby_by_string_property
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

const people = [
    { name: "Alice", dept: "eng" },
    { name: "Bob", dept: "eng" },
    { name: "Carol", dept: "hr" },
    { name: "Dave", dept: "hr" }
];
const groups = Object.groupBy(people, p => p.dept);
__check(__line(groups.eng.length), "2");
__check(__line(groups.hr.length), "2");
__check(__line(groups.eng[0].name), "Alice");
