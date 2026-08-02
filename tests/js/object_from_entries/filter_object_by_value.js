// vybe-test: js/object_from_entries/filter_object_by_value
// origin: languages/js/tests/js/test_object_from_entries.rs

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

const scores = { alice: 95, bob: 72, charlie: 88, dave: 61 };
const passing = Object.fromEntries(
    Object.entries(scores).filter(([, score]) => score >= 80)
);
__check(__line(Object.keys(passing).sort().join(",")), "alice,charlie");
