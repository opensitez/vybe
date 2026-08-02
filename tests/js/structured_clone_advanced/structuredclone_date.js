// vybe-test: js/structured_clone_advanced/structuredclone_date
// origin: languages/js/tests/js/test_structured_clone_advanced.rs

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

const d = new Date(2024, 0, 15);
const clone = structuredClone(d);
__check(__line(clone instanceof Date), "true");
__check(__line(clone.getFullYear()), "2024");
