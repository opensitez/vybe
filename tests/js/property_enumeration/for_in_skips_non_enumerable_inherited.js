// vybe-test: js/property_enumeration/for_in_skips_non_enumerable_inherited
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

const obj = {};
// toString is non-enumerable on Object.prototype
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("toString"));
