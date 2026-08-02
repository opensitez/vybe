// vybe-test: js/object_literal_advanced/computed_symbol_key_method
// origin: languages/js/tests/js/test_object_literal_advanced.rs

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

const sym = Symbol("method");
const obj = {
    [sym]() { return "from symbol method"; }
};
__check(__line(obj[sym]()), "from symbol method");
