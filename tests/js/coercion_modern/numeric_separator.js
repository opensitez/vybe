// vybe-test: js/coercion_modern/numeric_separator
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let million = 1_000_000;
let hex = 0xFF_FF;
let binary = 0b1010_0001;
__check(__line(million), "1000000");
__check(__line(hex), "65535");
__check(__line(binary), "161");
