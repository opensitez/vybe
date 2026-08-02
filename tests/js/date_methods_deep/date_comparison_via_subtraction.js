// vybe-test: js/date_methods_deep/date_comparison_via_subtraction
// origin: languages/js/tests/js/test_date_methods_deep.rs

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

const earlier = new Date("2024-01-01");
const later = new Date("2024-12-31");
__check(__line(later - earlier > 0), "true");
__check(__line(later > earlier), "true");
