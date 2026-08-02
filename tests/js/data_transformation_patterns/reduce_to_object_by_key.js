// vybe-test: js/data_transformation_patterns/reduce_to_object_by_key
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

const users = [
    { id: 1, name: "Alice" },
    { id: 2, name: "Bob" },
    { id: 3, name: "Charlie" }
];
const indexed = users.reduce((acc, u) => ({ ...acc, [u.id]: u }), {});
__check(__line(indexed[1].name), "Alice");
__check(__line(indexed[3].name), "Charlie");
