// vybe-test: js/regex_comprehensive/regex_global_replace_function
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

const result = "hello world".replace(/(\w+)/g, (match, word) => word[0].toUpperCase() + word.slice(1));
__check(__line(result), "Hello World");
