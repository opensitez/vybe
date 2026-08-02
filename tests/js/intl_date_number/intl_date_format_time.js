// vybe-test: js/intl_date_number/intl_date_format_time
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

const fmt = new Intl.DateTimeFormat("en-US", {
    hour: "2-digit", minute: "2-digit", second: "2-digit",
    timeZone: "UTC", hour12: false
});
const d = new Date("2024-01-01T14:30:45.000Z");
const result = fmt.format(d);
__check(__line(typeof result), "string");
__check(__line(result.includes("14") || result.includes("30")), "true");
