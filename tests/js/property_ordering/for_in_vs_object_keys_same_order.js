// vybe-test: js/property_ordering/for_in_vs_object_keys_same_order
// origin: languages/js/tests/js/test_property_ordering.rs

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

const obj = { a: 1, b: 2, c: 3 };
const forInKeys = [];
for (const k in obj) forInKeys.push(k);
const objectKeys = Object.keys(obj);
console.log(forInKeys.join(",") === objectKeys.join(","));
