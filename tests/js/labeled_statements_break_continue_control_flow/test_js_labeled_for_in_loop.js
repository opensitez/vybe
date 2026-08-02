// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_for_in_loop
// origin: languages/js/tests/js/test_js_labeled_statements_break_continue_control_flow.rs

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
const res = [];
forInLabel: for (const k in obj) {
    if (k === "b") break forInLabel;
    res.push(k);
}
console.log(res.join(","));
