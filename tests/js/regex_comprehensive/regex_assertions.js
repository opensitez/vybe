// vybe-test: js/regex_comprehensive/regex_assertions
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

const wordStart = "cat concatenate catalog".match(/\bcat\b/g);
__check(__line(wordStart.length), "1");
const atLineEnd = "hello\nworld".match(/hello$/m);
__check(__line(atLineEnd !== null), "true");
