// vybe-test: js/control_flow_patterns/while_loop_with_complex_condition
// origin: languages/js/tests/js/test_control_flow_patterns.rs

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

let a = 1, b = 100, count = 0;
while (a < b) {
    a *= 2;
    b -= 10;
    count++;
}
console.log(count);
console.log(a >= b);
