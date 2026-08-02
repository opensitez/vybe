// vybe-test: js/do_while_while_for_loop_control_flow/test_js_for_of_loop_sparse_array_includes_undefined_hole
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

const seen = [];
for (const value of [1, , 3]) {
    seen.push(value === undefined ? "u" : String(value));
}
console.log(seen.join(","));
