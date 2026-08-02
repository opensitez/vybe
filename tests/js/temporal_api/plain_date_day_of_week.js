// vybe-test: js/temporal_api/plain_date_day_of_week
// origin: languages/js/tests/js/test_temporal_api.rs

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

// Jan 1, 2024 was a Monday; getDay() returns 0=Sun,1=Mon,...
const d = new Date(2024, 0, 1);
__check(__line(d.getDay()), "1"); // 1 = Monday
