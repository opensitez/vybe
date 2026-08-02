// vybe-test: js/regex_advanced/regex_named_groups
// origin: languages/js/tests/js/test_regex_advanced.rs

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

let re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
let m = re.exec("2024-03-15");
__check(__line(m.groups.year), "2024");
__check(__line(m.groups.month), "03");
__check(__line(m.groups.day), "15");
