// vybe-test: js/logical_assignment/test_logical_assignment_in_for_loop_update
// origin: languages/js/tests/js/test_logical_assignment.rs

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

let state = null;
let log = [];
for (let i = 0; i < 3; i++, state ??= i) {
    log.push(String(state));
}
console.log(log.join(","));
