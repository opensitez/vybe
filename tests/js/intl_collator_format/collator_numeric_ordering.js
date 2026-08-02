// vybe-test: js/intl_collator_format/collator_numeric_ordering
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

const files = ["file10", "file2", "file1"];
const sorted = files.sort(new Intl.Collator("en", { numeric: true }).compare);
__check(__line(sorted.join(",")), "file1,file2,file10");
