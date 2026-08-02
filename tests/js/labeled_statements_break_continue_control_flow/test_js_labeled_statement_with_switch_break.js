// vybe-test: js/labeled_statements_break_continue_control_flow/test_js_labeled_statement_with_switch_break
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

const res = [];
outerLoop: for (let i = 1; i <= 2; i++) {
    switch (i) {
        case 1:
            res.push("case1");
            break; // Breaks switch, stays in for loop
        case 2:
            res.push("case2");
            break outerLoop; // Breaks outer for loop!
    }
    res.push("afterSwitch");
}
console.log(res.join(","));
