// vybe-test: js/property_ordering/keys_skips_symbols
// origin: languages/js/tests/js/test_property_ordering.rs

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
const obj = { a: 1, [sym]: 2, b: 3 };
__check(__line(Object.keys(obj).join(",")), "a,b");
