// vybe-test: js/ecma_strings/string_trim_start_end
// origin: languages/js/tests/js/test_ecma_strings.rs

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

const s = "\t hello \n";
__check(__line(s.trimStart() === "hello \n"), "true");
__check(__line(s.trimEnd() === "\t hello"), "true");
__check(__line(s.trim()), "hello");
