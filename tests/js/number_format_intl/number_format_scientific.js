// vybe-test: js/number_format_intl/number_format_scientific
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

const fmt = new Intl.NumberFormat("en", { notation: "scientific" });
const result = fmt.format(123456789);
__check(__line(result.includes("E")), "true");
