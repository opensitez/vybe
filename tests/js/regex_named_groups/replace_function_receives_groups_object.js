// vybe-test: js/regex_named_groups/replace_function_receives_groups_object
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

const result = "John Smith".replace(
    /(?<first>\w+) (?<last>\w+)/,
    (_, first, last, _offset, _str, groups) => groups.last + ", " + groups.first
);
__check(__line(result), "Smith, John");
