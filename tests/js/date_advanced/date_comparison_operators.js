// vybe-test: js/date_advanced/date_comparison_operators
// origin: languages/js/tests/js/test_date_advanced.rs

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

const a = new Date("2024-01-01");
const b = new Date("2024-06-01");
__check(__line(a < b), "true");
__check(__line(b > a), "true");
