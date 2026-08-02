// vybe-test: js/intl_collator_format/collator_compare_returns_number
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

const col = new Intl.Collator("en");
const r1 = col.compare("apple", "banana");
const r2 = col.compare("banana", "apple");
const r3 = col.compare("same", "same");
__check(__line(r1 < 0), "true");
__check(__line(r2 > 0), "true");
__check(__line(r3 === 0), "true");
