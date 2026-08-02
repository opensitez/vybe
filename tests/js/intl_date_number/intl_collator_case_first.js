// vybe-test: js/intl_date_number/intl_collator_case_first
// origin: languages/js/tests/js/test_intl_date_number.rs

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

const words = ["banana", "Apple", "cherry", "AVOCADO"];
// Case-insensitive sort
const sorted = words.sort(new Intl.Collator("en", { sensitivity: "base" }).compare);
__check(__line(sorted[0].toLowerCase()), "apple");
__check(__line(sorted.length), "4");
