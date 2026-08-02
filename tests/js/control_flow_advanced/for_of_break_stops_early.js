// vybe-test: js/control_flow_advanced/for_of_break_stops_early
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

const result = [];
for (const x of [10, 20, 30, 40]) {
    if (x === 30) break;
    result.push(x);
}
console.log(result.join(","));
