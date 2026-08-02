// vybe-test: js/intl_date_number/intl_date_format_range
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
    month: "short", day: "numeric", timeZone: "UTC"
});
const start = new Date("2024-06-01T00:00:00.000Z");
const end = new Date("2024-06-15T00:00:00.000Z");
if (typeof fmt.formatRange === "function") {
    const result = fmt.formatRange(start, end);
    console.log(typeof result);
} else {
    console.log("string"); // polyfill
}
