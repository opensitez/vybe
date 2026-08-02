// vybe-test: js/object_descriptors/object_values_returns_own_enumerable_values
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

const obj = { a: 10, b: 20, c: 30 };
const vals = Object.values(obj).sort((a, b) => a - b);
__check(__line(vals.join(",")), "10,20,30");
