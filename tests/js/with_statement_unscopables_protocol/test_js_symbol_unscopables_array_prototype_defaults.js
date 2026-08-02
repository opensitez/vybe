// vybe-test: js/with_statement_unscopables_protocol/test_js_symbol_unscopables_array_prototype_defaults
// origin: languages/js/tests/js/test_js_with_statement_unscopables_protocol.rs

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

const unscopables = Array.prototype[Symbol.unscopables];
__check(__line(unscopables.find + "|" + unscopables.includes + "|" + unscopables.flat), "true|true|true");
