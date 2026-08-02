// vybe-test: js/intl_collator_format/relative_time_format_days
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

const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto" });
__check(__line(rtf.format(-1, "day")), "yesterday");
__check(__line(rtf.format(1, "day")), "tomorrow");
