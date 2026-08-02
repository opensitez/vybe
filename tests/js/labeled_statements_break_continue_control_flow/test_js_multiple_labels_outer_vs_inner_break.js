// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_multiple_labels_outer_vs_inner_break
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

const log = [];
outer: inner: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (i === 1) break outer;
        if (j === 1) break inner;
        log.push(`${i}:${j}`);
    }
}
console.log(log.join("|"));
