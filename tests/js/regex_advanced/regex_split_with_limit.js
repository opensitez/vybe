// vybe-test: js/regex_advanced/regex_split_with_limit
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

let s = "a-b-c-d-e";
let parts = s.split("-", 3);
__check(__line(parts.join(",")), "a,b,c");
