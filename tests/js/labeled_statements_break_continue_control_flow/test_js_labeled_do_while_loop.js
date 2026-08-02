// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_do_while_loop
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

let i = 0;
const res = [];
loopLabel: do {
    i++;
    if (i === 2) continue loopLabel;
    res.push(i);
} while (i < 3);
console.log(res.join(","));
