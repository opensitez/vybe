// vybe-test: js/control_flow_patterns/for_in_uses_object_enumeration_order
// origin: languages/js/tests/js/test_control_flow_patterns.rs

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

const obj = { a: 1, c: 2, b: 3 };
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.join(","));
