// vybe-test: js/chaining_patterns/string_method_chaining
// origin: languages/js/tests/js/test_chaining_patterns.rs

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

const result = "  Hello, World!  "
    .trim()
    .toLowerCase()
    .replace(",", "")
    .split(" ")
    .join("-");
__check(__line(result), "hello-world!");
