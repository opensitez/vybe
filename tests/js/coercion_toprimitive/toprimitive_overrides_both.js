// vybe-test: js/coercion_toprimitive/toprimitive_overrides_both
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

const obj = {
    [Symbol.toPrimitive](hint) {
        switch (hint) {
            case "number": return 42;
            case "string": return "forty-two";
            default: return true;
        }
    }
};
__check(__line(+obj), "42");
__check(__line(`${obj}`), "forty-two");
__check(__line(obj + ""), "true");
