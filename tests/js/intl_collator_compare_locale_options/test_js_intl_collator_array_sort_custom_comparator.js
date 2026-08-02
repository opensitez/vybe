// vybe-test: js/intl_collator_compare_locale_options/test_js_intl_collator_array_sort_custom_comparator
// origin: languages/js/tests/js/test_js_intl_collator_compare_locale_options.rs

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

const words = ["banana", "Apple", "cherry"];
words.sort(new Intl.Collator("en", { sensitivity: "base" }).compare);
__check(__line(words.join(",")), "Apple,banana,cherry");
