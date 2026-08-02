// vybe-test: js/symbol_advanced/symbol_not_in_for_in
// origin: languages/js/tests/js/test_symbol_advanced.rs

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
const obj = { visible: 1, [sym]: 2 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
console.log(keys.includes("hidden"));
