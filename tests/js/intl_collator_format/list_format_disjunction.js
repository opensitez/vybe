// vybe-test: js/intl_collator_format/list_format_disjunction
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

const fmt = new Intl.ListFormat("en-US", { type: "disjunction" });
const result = fmt.format(["cats", "dogs"]);
__check(__line(result.includes("or")), "true");
