// vybe-test: js/string_algorithms/title_case
// origin: languages/js/tests/js/test_string_algorithms.rs

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

const titleCase = s => s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
__check(__line(titleCase("hello world")), "Hello World");
__check(__line(titleCase("the quick brown fox")), "The Quick Brown Fox");
