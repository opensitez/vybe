// vybe-test: js/regex_named_groups/named_group_undefined_when_not_matched_optional
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

const re = /(?<a>\d+)?(?<b>[a-z]+)/;
const m = re.exec("hello");
__check(__line(m.groups.a), "undefined");
__check(__line(m.groups.b), "hello");
