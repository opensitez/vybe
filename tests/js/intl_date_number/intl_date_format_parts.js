// vybe-test: js/intl_date_number/intl_date_format_parts
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
    year: "numeric", month: "long", day: "numeric",
    timeZone: "UTC"
});
const d = new Date("2024-06-15T00:00:00.000Z");
const parts = fmt.formatToParts(d);
const types = parts.map(p => p.type).join(",");
__check(__line(types.includes("year")), "true");
__check(__line(types.includes("month")), "true");
__check(__line(types.includes("day")), "true");
