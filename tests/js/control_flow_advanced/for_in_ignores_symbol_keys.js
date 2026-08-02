// vybe-test: js/control_flow_advanced/for_in_ignores_symbol_keys
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

const key = Symbol("secret");
const obj = { a: 1 };
obj[key] = 2;
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("a"));
console.log(keys.includes("secret"));
console.log(keys.includes(key));
