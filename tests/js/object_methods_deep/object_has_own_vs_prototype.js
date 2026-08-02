// vybe-test: js/object_methods_deep/object_has_own_vs_prototype
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

const obj = Object.create({ inherited: true });
obj.own = true;
__check(__line(Object.hasOwn(obj, "own")), "true");
__check(__line(Object.hasOwn(obj, "inherited")), "false");
