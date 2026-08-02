// vybe-test: js/operators_deep/instanceof_symbol_hasinstance_truthiness
// origin: languages/js/tests/js/test_operators_deep.rs

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

class FalseInstanceof {
    static [Symbol.hasInstance]() { return ""; }
}
class TruthyInstanceof {
    static [Symbol.hasInstance]() { return 42; }
}
__check(__line({} instanceof FalseInstanceof), "false");
__check(__line({} instanceof TruthyInstanceof), "true");
