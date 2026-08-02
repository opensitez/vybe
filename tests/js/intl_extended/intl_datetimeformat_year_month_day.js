// vybe-test: js/intl_extended/intl_datetimeformat_year_month_day
// origin: languages/js/tests/js/test_intl_extended.rs

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

const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" });
const result = fmt.format(new Date(2024, 0, 15));
__check(__line(result.includes("2024")), "true");
__check(__line(result.includes("January")), "true");
