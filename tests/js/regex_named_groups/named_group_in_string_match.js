// vybe-test: js/regex_named_groups/named_group_in_string_match
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

const m = "2024-03-21".match(/(?<y>\d{4})-(?<m>\d{2})-(?<d>\d{2})/);
__check(__line(m.groups.y), "2024");
