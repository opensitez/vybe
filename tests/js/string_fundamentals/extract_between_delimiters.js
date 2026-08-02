// vybe-test: js/string_fundamentals/extract_between_delimiters
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

const s = "prefix[content]suffix";
const start = s.indexOf("[") + 1;
const end = s.indexOf("]");
__check(__line(s.slice(start, end)), "content");
