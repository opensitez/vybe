// vybe-test: js/object_introspection/object_from_entries_from_array
// origin: languages/js/tests/js/test_object_introspection.rs

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

const pairs = [["name", "Bob"], ["age", 25], ["city", "Paris"]];
const obj = Object.fromEntries(pairs);
__check(__line(obj.name), "Bob");
__check(__line(obj.age), "25");
__check(__line(obj.city), "Paris");
