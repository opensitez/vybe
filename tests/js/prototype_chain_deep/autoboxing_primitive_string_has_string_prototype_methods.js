// vybe-test: js/prototype_chain_deep/autoboxing_primitive_string_has_string_prototype_methods
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

const s = "hello";
__check(__line(s.toUpperCase()), "HELLO");
__check(__line(s.length), "5");
