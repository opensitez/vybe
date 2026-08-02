// vybe-test: js/class_inheritance_advanced/extends_null_sets_prototype_to_null
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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

class NullPrototype extends null {}
__check(__line(Object.getPrototypeOf(NullPrototype.prototype) === null), "false");
__check(__line(Object.getPrototypeOf(NullPrototype) === Function.prototype), "true");
