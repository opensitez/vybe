// vybe-test: js/string_processing_deep/number_formatting_custom
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function formatNum(n, decimals = 2, sep = ",") {
    const [int, dec] = n.toFixed(decimals).split(".");
    const formatted = int.replace(/\B(?=(\d{3})+(?!\d))/g, sep);
    return dec ? `${formatted}.${dec}` : formatted;
}
__check(__line(formatNum(1234567.89)), "1,234,567.89");
__check(__line(formatNum(1000, 0)), "1,000");
__check(__line(formatNum(42.1234, 3)), "42.123");
