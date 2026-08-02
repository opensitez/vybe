// vybe-test: js/do_while_while_for_loop_control_flow/test_js_nested_loops_with_labeled_break
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
outer: for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 0) continue;
        if (j === 2) break outer;
        seen.push(i + ":" + j);
    }
}
console.log(seen.join(","));
