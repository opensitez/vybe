// vybe-test: js/object_descriptors/define_property_non_enumerable_hidden_from_for_in
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

const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false, writable: true, configurable: true });
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("visible"));
console.log(keys.includes("hidden"));
