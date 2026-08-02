// vybe-test: js/intl_date_number/intl_date_format_epoch_number
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

const fmt = new Intl.DateTimeFormat("en", { timeZone: "UTC" });
__check(__line(fmt.format(0).includes("1970")), "true");
