// vybe-test: js/regex_named_groups/replace_uses_named_group_in_template
// origin: languages/js/tests/js/test_regex_named_groups.rs

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

const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
const result = "2024-01-15".replace(re, "$<day>/$<month>/$<year>");
__check(__line(result), "15/01/2024");
