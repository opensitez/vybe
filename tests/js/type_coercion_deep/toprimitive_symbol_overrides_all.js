// vybe-test: js/type_coercion_deep/toprimitive_symbol_overrides_all
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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
        if (hint === "number") return 10;
        if (hint === "string") return "ten";
        return true; // default
    }
};
__check(__line(+obj), "10");
__check(__line(`${obj}`), "ten");
__check(__line(obj + ""), "true");
