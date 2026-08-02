// vybe-test: js/string_methods_more/string_localecompare_order
// origin: languages/js/tests/js/test_string_methods_more.rs

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

const words = ["banana", "apple", "cherry"];
words.sort((a, b) => a < b ? -1 : a > b ? 1 : 0);
__check(__line(words.join(",")), "apple,banana,cherry");
