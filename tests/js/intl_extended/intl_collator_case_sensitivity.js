// vybe-test: js/intl_extended/intl_collator_case_sensitivity
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

const coll = new Intl.Collator("en-US", { sensitivity: "base" });
__check(__line(coll.compare("a", "A") === 0), "true");
