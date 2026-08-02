// vybe-test: js/delete_operator/delete_non_configurable_returns_false_in_sloppy
// origin: languages/js/tests/js/test_delete_operator.rs

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

const obj = {};
Object.defineProperty(obj, "fixed", { value: 1, configurable: false });
const result = delete obj.fixed;
__check(__line(result), "false");
__check(__line(obj.fixed), "1");
