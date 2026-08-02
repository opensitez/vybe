// vybe-test: js/regex_flags_advanced/regex_source_and_flags_properties
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const re = /hello\s+world/gi;
__check(__line(re.source), "hello\\s+world");
__check(__line(re.flags), "gi");
