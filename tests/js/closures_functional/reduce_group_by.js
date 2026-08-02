// vybe-test: js/closures_functional/reduce_group_by
// origin: languages/js/tests/js/test_closures_functional.rs

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

let people = [
    { name: "Alice", dept: "eng" },
    { name: "Bob", dept: "sales" },
    { name: "Charlie", dept: "eng" },
    { name: "Diana", dept: "sales" },
    { name: "Eve", dept: "eng" }
];
let groups = people.reduce((acc, p) => {
    if (!acc[p.dept]) acc[p.dept] = [];
    acc[p.dept].push(p.name);
    return acc;
}, {});
__check(__line(groups.eng.length), "3");
__check(__line(groups.sales.length), "2");
