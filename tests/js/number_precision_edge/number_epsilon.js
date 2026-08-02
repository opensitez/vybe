// vybe-test: js/number_precision_edge/number_epsilon
// origin: languages/js/tests/js/test_number_precision_edge.rs

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

console.log(Number.EPSILON > 0);
console.log(Number.EPSILON < 0.001);
// Using epsilon for float comparison
const a = 0.1 + 0.2;
const b = 0.3;
console.log(Math.abs(a - b) < Number.EPSILON * 10);
