// vybe-test: js/do_while_while_for_loop_control_flow/test_js_for_in_loop_skips_symbols
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

const sym = Symbol("id");
const obj = { stringProp: 1, [sym]: 2 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(","));
