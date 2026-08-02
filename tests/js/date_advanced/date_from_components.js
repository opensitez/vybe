// vybe-test: js/date_advanced/date_from_components
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

const d = new Date(2024, 0, 15); // January (0-indexed)
__check(__line(d.getFullYear()), "2024");
__check(__line(d.getMonth()), "0");
__check(__line(d.getDate()), "15");
