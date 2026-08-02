// vybe-test: js/number_format_intl/number_format_fraction_digits
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

const fmt = new Intl.NumberFormat("en", { minimumFractionDigits: 2, maximumFractionDigits: 4 });
__check(__line(fmt.format(3.14159)), "3.1416");
__check(__line(fmt.format(1)), "1.00");
