// vybe-test: js/date/date_setters
// origin: languages/js/tests/js/test_date.rs

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

let d = new Date(2024, 0, 1);
d.setFullYear(2025);
d.setMonth(11);
d.setDate(25);
__check(__line(d.getFullYear()), "2025");
__check(__line(d.getMonth()), "11");
__check(__line(d.getDate()), "25");
