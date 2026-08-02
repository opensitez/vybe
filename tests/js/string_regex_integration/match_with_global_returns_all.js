// vybe-test: js/string_regex_integration/match_with_global_returns_all
// origin: languages/js/tests/js/test_string_regex_integration.rs

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

const matches = "test1 test2 test3".match(/test\d/g);
__check(__line(matches.join(",")), "test1,test2,test3");
