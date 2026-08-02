// vybe-test: js/date/date_set_hours_minutes
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

let d = new Date(2024, 0, 1, 0, 0, 0);
d.setHours(14);
d.setMinutes(30);
d.setSeconds(59);
__check(__line(d.getHours()), "14");
__check(__line(d.getMinutes()), "30");
__check(__line(d.getSeconds()), "59");
