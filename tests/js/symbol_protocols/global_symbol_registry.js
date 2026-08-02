// vybe-test: js/symbol_protocols/global_symbol_registry
// origin: languages/js/tests/js/test_symbol_protocols.rs

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

const s1 = Symbol.for("shared");
const s2 = Symbol.for("shared");
console.log(s1 === s2);
console.log(Symbol.keyFor(s1));
const local = Symbol("local");
console.log(Symbol.keyFor(local));
