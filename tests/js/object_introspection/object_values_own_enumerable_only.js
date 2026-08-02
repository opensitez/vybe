// vybe-test: js/object_introspection/object_values_own_enumerable_only
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

const obj = { p: 10, q: 20 };
Object.defineProperty(obj, "secret", { value: 99, enumerable: false });
const vals = Object.values(obj);
__check(__line(vals.join(",")), "10,20");
__check(__line(vals.includes(99)), "false");
