// vybe-test: js/edge_cases_final/comma_operator_in_for
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

let a = 0, b = 10;
for (let i = 0; i < 5; i++, b--) { a += i; }
console.log(a);
console.log(b);
