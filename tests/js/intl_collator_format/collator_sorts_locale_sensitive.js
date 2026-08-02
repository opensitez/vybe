// vybe-test: js/intl_collator_format/collator_sorts_locale_sensitive
// origin: languages/js/tests/js/test_intl_collator_format.rs

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

const words = ["Zebra", "apple", "Banana"];
const sorted = words.sort(new Intl.Collator("en", { sensitivity: "base" }).compare);
// Case-insensitive sort: apple, Banana, Zebra
__check(__line(sorted[0].toLowerCase()), "apple");
