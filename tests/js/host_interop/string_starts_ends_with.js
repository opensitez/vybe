// vybe-test: js/host_interop/string_starts_ends_with
// origin: languages/js/tests/js/js_host_interop_test.rs

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

__check(__line("hello".startsWith("hel")), "true");
        __check(__line("hello".endsWith("llo")), "true");
        __check(__line("hello".startsWith("xyz")), "false");
