// vybe-test: js/string_processing_patterns/truncate_with_ellipsis
// origin: languages/js/tests/js/test_string_processing_patterns.rs

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

function truncate(str, max, ellipsis = "...") {
    if (str.length <= max) return str;
    return str.slice(0, max - ellipsis.length) + ellipsis;
}
__check(__line(truncate("Hello, World!", 8)), "Hello...");
__check(__line(truncate("Short", 10)), "Short");
