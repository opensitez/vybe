// vybe-test: js/intl_collator_format/number_format_notation_compact
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

const fmt = new Intl.NumberFormat("en-US", { notation: "compact" });
const result = fmt.format(1000000);
// Should be "1M" or similar
__check(__line(result.includes("M") || result.includes("m") || result.length < 8), "true");
