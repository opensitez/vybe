// vybe-test: js/property_enumeration/for_in_skips_symbol_keys
// origin: languages/js/tests/js/test_property_enumeration.rs

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
const obj = { a: 1 };
obj[sym] = 2;
const keys = [];
for (const key in obj) {
    keys.push(key);
}
console.log(keys.includes("a"));
console.log(keys.includes("s"));
console.log(keys.includes(sym));
console.log(Object.keys(obj).includes("a"));
