// vybe-test: js/symbol_wellknown/symbol_not_in_json_stringify
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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

const sym = Symbol("hidden");
const obj = { [sym]: "secret", visible: "yes" };
const json = JSON.stringify(obj);
__check(__line(json), "{\"visible\":\"yes\"}");
