// vybe-test: js/intl_collator_format/datetimeformat_weekday_option
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

// June 10 2024 is a Monday
const date = new Date(2024, 5, 10);
const fmt = new Intl.DateTimeFormat("en-US", { weekday: "long" });
const result = fmt.format(date);
__check(__line(result === "Monday"), "true");
