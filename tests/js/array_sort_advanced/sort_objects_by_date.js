// vybe-test: js/array_sort_advanced/sort_objects_by_date
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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

const events = [
    { name: "c", date: new Date(2024, 2, 1) },
    { name: "a", date: new Date(2024, 0, 1) },
    { name: "b", date: new Date(2024, 1, 1) },
];
events.sort((a, b) => a.date - b.date);
__check(__line(events.map(e => e.name).join(",")), "a,b,c");
