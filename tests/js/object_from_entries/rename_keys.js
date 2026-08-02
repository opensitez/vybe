// vybe-test: js/object_from_entries/rename_keys
// origin: languages/js/tests/js/test_object_from_entries.rs

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

const keyMap = { firstName: "first_name", lastName: "last_name" };
const user = { firstName: "Alice", lastName: "Smith" };
const renamed = Object.fromEntries(
    Object.entries(user).map(([k, v]) => [keyMap[k] || k, v])
);
__check(__line(renamed.first_name), "Alice");
__check(__line(renamed.last_name), "Smith");
