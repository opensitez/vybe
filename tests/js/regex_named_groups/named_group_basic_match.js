// vybe-test: js/regex_named_groups/named_group_basic_match
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
const m = re.exec("2024-01-15");
__check(__line(m.groups.year), "2024");
__check(__line(m.groups.month), "01");
__check(__line(m.groups.day), "15");
