// vybe-test: js/data_transformation_patterns/sort_by_multiple_keys
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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
    { name: "Bob", age: 30 },
    { name: "Alice", age: 25 },
    { name: "Charlie", age: 30 },
    { name: "Alice", age: 20 },
];
people.sort((a, b) => {
    if (a.name !== b.name) return a.name.localeCompare(b.name);
    return a.age - b.age;
});
__check(__line(people.map(p => p.name + p.age).join(",")), "Alice20,Alice25,Bob30,Charlie30");
