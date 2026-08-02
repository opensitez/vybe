// vybe-test: js/control_flow_advanced/for_of_continue_and_finally_in_loop
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

const out = [];
for (const item of [1, 2, 3]) {
    try {
        if (item === 2) {
            continue;
        }
        out.push("body-" + item);
    } finally {
        out.push("finally-" + item);
    }
}
console.log(out.join("|"));
