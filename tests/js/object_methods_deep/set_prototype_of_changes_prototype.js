// vybe-test: js/object_methods_deep/set_prototype_of_changes_prototype
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

const base = { base: true };
const obj = {};
const result = Object.setPrototypeOf(obj, base);
__check(__line(result === obj), "true");
__check(__line(Object.getPrototypeOf(obj) === base), "true");
__check(__line(obj.base), "true");
