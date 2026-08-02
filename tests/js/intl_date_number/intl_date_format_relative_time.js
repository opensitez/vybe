// vybe-test: js/intl_date_number/intl_date_format_relative_time
// origin: languages/js/tests/js/test_intl_date_number.rs

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

const fmt = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
__check(__line(fmt.format(-1, "day")), "yesterday");
__check(__line(fmt.format(1, "day")), "tomorrow");
__check(__line(fmt.format(-7, "week")), "7 weeks ago");
