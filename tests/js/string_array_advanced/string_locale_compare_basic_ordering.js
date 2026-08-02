// vybe-test: js/string_array_advanced/string_locale_compare_basic_ordering
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

__check(__line("a".localeCompare("b") < 0), "true");
__check(__line("b".localeCompare("a") > 0), "true");
__check(__line("a".localeCompare("a") === 0), "true");
