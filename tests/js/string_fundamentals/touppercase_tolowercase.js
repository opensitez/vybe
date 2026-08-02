// vybe-test: js/string_fundamentals/touppercase_tolowercase
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

__check(__line("Hello World".toUpperCase()), "HELLO WORLD");
__check(__line("Hello World".toLowerCase()), "hello world");
