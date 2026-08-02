// vybe-test: js/regex_advanced/regex_constructor
// origin: languages/js/tests/js/test_regex_advanced.rs

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

let pattern = "hello";
let re = new RegExp(pattern, "i");
__check(__line(re.test("Hello World")), "true");
__check(__line(re.test("hi")), "false");
