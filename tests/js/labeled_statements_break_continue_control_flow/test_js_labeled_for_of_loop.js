// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_for_of_loop
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

const arr = [10, 20, 30];
const res = [];
forOfLabel: for (const val of arr) {
    if (val === 20) continue forOfLabel;
    res.push(val);
}
console.log(res.join(","));
