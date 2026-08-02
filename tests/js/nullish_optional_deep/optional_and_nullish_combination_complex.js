// vybe-test: js/nullish_optional_deep/optional_and_nullish_combination_complex
// origin: languages/js/tests/js/test_nullish_optional_deep.rs

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

const config = {
    settings: null,
    timeout: 0,
};
const timeout = config.settings?.timeout ?? config.timeout ?? 3000;
__check(__line(timeout), "0");
