// vybe-test: js/object_introspection/define_property_non_enumerable_absent_from_for_in
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

const obj = { a: 1 };
Object.defineProperty(obj, "secret", { value: 42, enumerable: false });
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("a"));
console.log(keys.includes("secret"));
console.log(obj.secret);
