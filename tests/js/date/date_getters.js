// vybe-test: js/date/date_getters
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

let d = new Date(2024, 5, 15, 10, 30, 45);
__check(__line(d.getFullYear()), "2024");
__check(__line(d.getMonth()), "5");
__check(__line(d.getDate()), "15");
__check(__line(d.getHours()), "10");
__check(__line(d.getMinutes()), "30");
__check(__line(d.getSeconds()), "45");
