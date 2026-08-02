// vybe-test: js/do_while_while_for_loop_control_flow/test_js_for_in_loop_deleting_property_during_enumeration
// origin: languages/js/tests/js/test_js_do_while_while_for_loop_control_flow.rs

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
for (const k in obj) {
    keys.push(k);
    if (k === "a") delete obj.b; // Deleting 'b' before visited prevents enumeration!
}
console.log(keys.join(","));
