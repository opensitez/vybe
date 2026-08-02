// vybe-test: js/control_flow_advanced/test_js_control_flow_do_while_condition_side_effects
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

let bodyCount = 0;
let condCount = 0;
do {
    bodyCount++;
} while ((condCount += 5) < 15);
console.log(`${bodyCount}|${condCount}`);
