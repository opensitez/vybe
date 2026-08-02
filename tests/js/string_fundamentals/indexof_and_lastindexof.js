// vybe-test: js/string_fundamentals/indexof_and_lastindexof
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

const s = "abcabc";
__check(__line(s.indexOf("b")), "1");
__check(__line(s.lastIndexOf("b")), "4");
__check(__line(s.indexOf("x")), "-1");
