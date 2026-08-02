// vybe-test: js/object_descriptors/integer_indices_come_before_string_keys_in_enumeration
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const obj = { b: 2, a: 1, "0": "zero", "2": "two", "1": "one" };
const keys = Object.keys(obj);
__check(__line(keys.includes("0")), "true");
__check(__line(keys.includes("a")), "true");
__check(__line(keys.length), "5");
