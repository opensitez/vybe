// vybe-test: js/edge_cases_final/nullish_assign_short_circuit
// origin: languages/js/tests/js/test_edge_cases_final.rs

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

let calls = 0;
const expensive = () => ++calls;
let a = "existing";
a ??= expensive();
__check(__line(a), "existing");
__check(__line(calls), "0");
let b = null;
b ??= expensive();
__check(__line(b), "1");
__check(__line(calls), "1");
