// vybe-test: js/new_globals_e2e/symbol_for_returns_same_value
// origin: languages/js/tests/js/test_new_globals_e2e.rs

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

const a = Symbol.for("shared-key");
        const b = Symbol.for("shared-key");
        console.log(String(a) === String(b));
