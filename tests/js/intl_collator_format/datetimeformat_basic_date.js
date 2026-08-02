// vybe-test: js/intl_collator_format/datetimeformat_basic_date
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

const date = new Date(2024, 0, 15); // Jan 15, 2024
const fmt = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "numeric" });
const result = fmt.format(date);
__check(__line(result.includes("2024")), "true");
__check(__line(result.includes("January") || result.includes("Jan")), "true");
