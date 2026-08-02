// vybe-test: js/object_spread_edge/spread_copies_symbol_keys
// origin: languages/js/tests/js/test_object_spread_edge.rs

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

const sym = Symbol("s");
const src = { [sym]: 42, str: "ok" };
const result = { ...src };
__check(__line(result[sym]), "42");
