// vybe-test: js/for_in_deep/for_in_with_symbol_key_skips
// origin: languages/js/tests/js/test_for_in_deep.rs

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
const obj = { a: 1, [sym]: 2 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
