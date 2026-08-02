// vybe-test: js/object_literal_advanced/computed_property_from_variable
// origin: languages/js/tests/js/test_object_literal_advanced.rs

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

const key = "name";
const obj = { [key]: "Alice" };
__check(__line(obj.name), "Alice");
__check(__line(obj[key]), "Alice");
