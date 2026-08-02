// vybe-test: js/string_methods_deep/match_without_global_flag
// origin: languages/js/tests/js/test_string_methods_deep.rs

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

const match = "abc123def456".match(/\d+/);
__check(__line(match[0]), "123");
__check(__line(match.index), "3");
