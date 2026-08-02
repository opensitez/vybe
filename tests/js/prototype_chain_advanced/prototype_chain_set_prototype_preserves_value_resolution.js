// vybe-test: js/prototype_chain_advanced/prototype_chain_set_prototype_preserves_value_resolution
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

const base = { label() { return "base"; } };
const replacement = { label() { return "replacement"; } };
const obj = Object.create(base);
__check(__line(obj.label()), "base");
Object.setPrototypeOf(obj, replacement);
__check(__line(obj.label()), "replacement");
__check(__line(Object.getPrototypeOf(obj) === replacement), "true");
