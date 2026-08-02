// vybe-test: js/do_while_while_for_loop_control_flow/test_js_labeled_do_while_with_continue_and_break
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

let out = [];
let i = 0;
outer: do {
    i++;
    if (i === 2) {
        out.push("continue-" + i);
        continue outer;
    }
    if (i === 4) {
        out.push("break-" + i);
        break;
    }
    out.push("body-" + i);
} while (i < 5);
console.log(out.join("|"));
