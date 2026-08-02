// vybe-test: js/control_flow_advanced/comma_operator_in_for_update
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

let sum = 0;
for (let i = 0, j = 10; i < 3; i++, j--) {
    sum += j;
}
console.log(sum);
