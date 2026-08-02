// vybe-test: js/object_methods_deep/object_keys_only_own_enumerable
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Object.keys(obj);
__check(__line(keys.sort().join(",")), "a,b");
