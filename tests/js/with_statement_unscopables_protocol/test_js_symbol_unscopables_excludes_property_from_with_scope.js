// vybe-test: js/with_statement_unscopables_protocol/test_js_symbol_unscopables_excludes_property_from_with_scope
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

const a = "OuterA";
const obj = {
    a: "InnerA",
    [Symbol.unscopables]: {
        a: true // Excludes 'a' from being bound in with(obj) scope!
    }
};
with (obj) {
    __check(__line(a), "OuterA");
}
