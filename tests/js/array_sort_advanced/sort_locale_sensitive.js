// vybe-test: js/array_sort_advanced/sort_locale_sensitive
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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

const words = ["résumé", "apple", "éclair"];
words.sort((a, b) => a.localeCompare(b));
// All should be sorted — exact order depends on locale but shouldn't throw
console.log(words.length);
