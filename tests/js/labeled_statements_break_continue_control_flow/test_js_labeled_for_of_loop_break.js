// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_for_of_loop_break
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

const arr = [1, 2, 3];
const res = [];
outer: for (const x of arr) {
    if (x === 2) break outer;
    res.push(x);
}
console.log(res.join(","));
