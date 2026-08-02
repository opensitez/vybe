// vybe-test: js/string_fundamentals/string_split_with_limit_and_empty_pattern
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

__check(__line("a,b,c".split(",", 2).join("|")), "a|b");
__check(__line("abc".split("").join("-")), "a-b-c");
