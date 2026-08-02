// vybe-test: js/regex_named_groups/sticky_flag_matches_from_lastindex
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

const re = /\d+/y;
re.lastIndex = 2;
const m = re.exec("ab123");
__check(__line(m[0]), "123");
__check(__line(re.lastIndex), "5");
