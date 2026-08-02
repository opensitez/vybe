// vybe-test: js/property_enumeration/object_entries_follows_key_order
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const obj = { z: 3, a: 1, m: 2 };
const entries = Object.entries(obj);
// Insertion order for non-integer keys
console.log(entries.map(([k]) => k).join(","));
