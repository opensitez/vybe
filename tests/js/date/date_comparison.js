// vybe-test: js/date/date_comparison
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

let d1 = new Date(2024, 0, 1);
let d2 = new Date(2024, 6, 1);
__check(__line(d1 < d2), "true");
__check(__line(d1 > d2), "false");
__check(__line(d1.getTime() === new Date(2024, 0, 1).getTime()), "true");
