// vybe-test: js/spread_rest_advanced/spread_in_new_constructor
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

const args = [2024, 0, 1];
const d = new Date(...args);
__check(__line(d.getFullYear()), "2024");
