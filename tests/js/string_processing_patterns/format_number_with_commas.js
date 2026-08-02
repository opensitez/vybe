// vybe-test: js/string_processing_patterns/format_number_with_commas
// origin: languages/js/tests/js/test_string_processing_patterns.rs

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

function formatNumber(n) {
    return n.toLocaleString("en-US");
}
// Or use regex:
function formatNumberRegex(n) {
    return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
__check(__line(formatNumberRegex(1234567)), "1,234,567");
__check(__line(formatNumberRegex(1000)), "1,000");
__check(__line(formatNumberRegex(42)), "42");
