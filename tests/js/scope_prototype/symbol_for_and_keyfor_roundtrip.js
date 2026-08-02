// vybe-test: js/scope_prototype/symbol_for_and_keyfor_roundtrip
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let sym = Symbol.for("shared.key");
console.log(Symbol.keyFor(sym));
console.log(Symbol.keyFor(Symbol("local")));
