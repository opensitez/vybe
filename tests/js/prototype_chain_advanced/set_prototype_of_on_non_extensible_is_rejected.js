// vybe-test: js/prototype_chain_advanced/set_prototype_of_on_non_extensible_is_rejected
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

const obj = {};
Object.preventExtensions(obj);
try {
    Object.setPrototypeOf(obj, { tag: "next" });
    console.log("changed");
} catch (e) {
    console.log(e.name);
}
