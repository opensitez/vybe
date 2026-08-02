// vybe-test: js/coercion_modern/optional_chaining_all_forms
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

let obj = {
    a: { b: { c: 42 } },
    fn: () => "called"
};
__check(__line(obj?.a?.b?.c), "42");
__check(__line(obj?.x?.y?.z), "undefined");
__check(__line(obj?.fn?.()), "called");
__check(__line(obj?.missing?.()), "undefined");
