// vybe-test: js/symbol_wellknown/symbol_unique_per_call
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

const s1 = Symbol("id");
const s2 = Symbol("id");
__check(__line(s1 === s2), "false");
__check(__line(typeof s1), "symbol");
