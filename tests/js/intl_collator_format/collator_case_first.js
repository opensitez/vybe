// vybe-test: js/intl_collator_format/collator_case_first
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

const col = new Intl.Collator("en", { caseFirst: "upper" });
const result = col.compare("A", "a");
__check(__line(typeof result === "number"), "true");
