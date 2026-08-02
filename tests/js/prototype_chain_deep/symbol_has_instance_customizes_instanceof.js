// vybe-test: js/prototype_chain_deep/symbol_has_instance_customizes_instanceof
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

class EvenCheck {
    static [Symbol.hasInstance](val) {
        return typeof val === "number" && val % 2 === 0;
    }
}
__check(__line(2 instanceof EvenCheck), "true");
__check(__line(3 instanceof EvenCheck), "false");
