// vybe-test: js/number_format_intl/test_number_format_formattoparts_is_function
// origin: languages/js/tests/js/test_number_format_intl.rs

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

const fmt = new Intl.NumberFormat("en");
__check(__line(typeof fmt.formatToParts === "function"), "true");
