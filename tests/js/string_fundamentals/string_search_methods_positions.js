// vybe-test: js/string_fundamentals/string_search_methods_positions
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

__check(__line("javascript".startsWith("java")), "true");
__check(__line("javascript".startsWith("ava", 1)), "true");
__check(__line("javascript".endsWith("ipt")), "true");
__check(__line("javascript".endsWith("java", 4)), "true");
__check(__line("javascript".includes("script")), "true");
__check(__line("javascript".includes("java", 1)), "false");
