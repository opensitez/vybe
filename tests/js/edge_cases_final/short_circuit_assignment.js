// vybe-test: js/edge_cases_final/short_circuit_assignment
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

let count = 0;
const inc = () => ++count;
false && inc();
true || inc();
null ?? inc();
__check(__line(count), "1");
true && inc();
false || inc();
"something" ?? inc();
__check(__line(count), "3");
