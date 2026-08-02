// vybe-test: js/prototype_chain_deep/string_instances_inherit_from_string_prototype
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

const s = new String("hello");
__check(__line(s instanceof String), "true");
__check(__line(Object.getPrototypeOf(s) === String.prototype), "true");
