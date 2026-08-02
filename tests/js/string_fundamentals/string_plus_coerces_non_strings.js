// vybe-test: js/string_fundamentals/string_plus_coerces_non_strings
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

__check(__line("value: " + 42), "value: 42");
__check(__line("flag: " + true), "flag: true");
__check(__line("nothing: " + null), "nothing: null");
