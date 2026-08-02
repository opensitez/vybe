// vybe-test: js/coercion_toprimitive/array_plus_behaviors
// origin: languages/js/tests/js/test_coercion_toprimitive.rs

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

__check(__line([] + []), "");
__check(__line([] + {}), "[object Object]");
__check(__line({} + []), "[object Object]");
__check(__line([1] + [2]), "12");
