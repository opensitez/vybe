// vybe-test: js/control_flow_advanced/finally_always_runs_on_normal_exit
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

let log = [];
function f() {
    try { log.push("try"); return 1; }
    finally { log.push("finally"); }
}
f();
__check(__line(log.join(",")), "try,finally");
