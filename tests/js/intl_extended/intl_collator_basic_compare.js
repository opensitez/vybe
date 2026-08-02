// vybe-test: js/intl_extended/intl_collator_basic_compare
// origin: languages/js/tests/js/test_intl_extended.rs

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

const coll = new Intl.Collator("en-US");
const words = ["banana", "apple", "cherry"];
words.sort(coll.compare);
__check(__line(words[0]), "apple");
