// vybe-test: js/intl_extended/intl_relativetimeformat_basic
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

const rtf = new Intl.RelativeTimeFormat("en-US", { numeric: "auto" });
const result = rtf.format(-1, "day");
__check(__line(result), "yesterday");
