// vybe-test: js/regex_comprehensive/named_capture_groups_replace
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

const date = "2024-06-15";
const reordered = date.replace(
    /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/,
    "$<day>/$<month>/$<year>"
);
__check(__line(reordered), "15/06/2024");
