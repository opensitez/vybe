// vybe-test: js/function_edge_cases/arguments_spread_to_array
// origin: languages/js/tests/js/test_function_edge_cases.rs

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

function toArr() { return Array.from(arguments); }
const a = toArr(10, 20, 30);
__check(__line(a.join(",")), "10,20,30");
