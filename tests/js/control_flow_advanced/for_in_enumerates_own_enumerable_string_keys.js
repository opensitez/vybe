// vybe-test: js/control_flow_advanced/for_in_enumerates_own_enumerable_string_keys
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

const obj = { a: 1, b: 2, c: 3 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.sort().join(","));
