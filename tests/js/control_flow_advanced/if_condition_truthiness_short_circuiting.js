// vybe-test: js/control_flow_advanced/if_condition_truthiness_short_circuiting
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

let hit = 0;
if ("") {
    hit += 1;
} else if (0 || false) {
    hit += 10;
} else {
    hit += 100;
}
__check(__line(hit), "100");
