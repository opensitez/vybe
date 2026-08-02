// vybe-test: js/regex_comprehensive/negative_lookahead_pattern
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

// Match words not followed by a comma. "world" is followed by "," and is excluded.
const text = "hello, world, foo bar baz";
const words = text.match(/\b\w+\b(?!,)/g);
__check(__line(words.join(",")), "foo,bar,baz");
