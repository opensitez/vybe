// vybe-test: js/regex_string_methods/match_without_global_returns_first_with_groups
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const result = "hello world".match(/(\w+)\s(\w+)/);
__check(__line(result[0]), "hello world");
__check(__line(result[1]), "hello");
__check(__line(result[2]), "world");
__check(__line(result.index), "0");
