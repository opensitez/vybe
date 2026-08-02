// vybe-test: js/intl_e2e/intl_number_format_default_decimal
// origin: languages/js/tests/js/test_intl_e2e.rs

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

const nf = new Intl.NumberFormat();
        __check(__line(nf.format(1234.5)), "1,234.5");
