// vybe-test: js/symbol_for_key_for_registry/test_js_symbol_for_coerces_key_argument_to_string
// origin: languages/js/tests/js/test_js_symbol_for_key_for_registry.rs

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

const s1 = Symbol.for(100);
const s2 = Symbol.for("100");
console.log(s1 === s2 + "|" + Symbol.keyFor(s1));
