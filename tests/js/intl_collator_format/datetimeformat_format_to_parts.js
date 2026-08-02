// vybe-test: js/intl_collator_format/datetimeformat_format_to_parts
// origin: languages/js/tests/js/test_intl_collator_format.rs

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

const date = new Date(2024, 5, 10);
const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "2-digit", day: "2-digit" });
const parts = fmt.formatToParts(date);
const types = parts.map(p => p.type);
__check(__line(types.includes("year")), "true");
__check(__line(types.includes("month")), "true");
__check(__line(types.includes("day")), "true");
