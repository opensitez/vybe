// vybe-test: js/control_flow_advanced/do_while_loops_until_condition_false
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

let n = 1;
let product = 1;
do {
    product *= n;
    n++;
} while (n <= 5);
console.log(product);
