// vybe-test: js/prototype_chain_deep/set_prototype_to_null_removes_object_prototype_methods
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

const obj = { value: 1 };
Object.setPrototypeOf(obj, null);
__check(__line(Object.getPrototypeOf(obj) === null), "true");
__check(__line(typeof obj.toString), "undefined");
