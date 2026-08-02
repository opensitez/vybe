// vybe-test: js/coercion_toprimitive/test_toprimitive_returning_object_throws_typeerror
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

const bad = {
    [Symbol.toPrimitive]() {
        return {};
    }
};
try {
    +bad;
} catch (e) {
    __check(__line(e.name), "TypeError");
}
