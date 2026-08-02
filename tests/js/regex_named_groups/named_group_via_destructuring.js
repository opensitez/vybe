// vybe-test: js/regex_named_groups/named_group_via_destructuring
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

const { groups: { first, last } } = /(?<first>\w+) (?<last>\w+)/.exec("John Doe");
__check(__line(first), "John");
__check(__line(last), "Doe");
