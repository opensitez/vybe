// vybe-test: js/reflect_api/reflect_delete_non_configurable_fails
// origin: languages/js/tests/js/test_reflect_api.rs

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
Object.defineProperty(obj, "x", { value: 1, configurable: false });
__check(__line(Reflect.deleteProperty(obj, "x")), "false"); // false
__check(__line("x" in obj), "true");
