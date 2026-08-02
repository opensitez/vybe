// vybe-test: js/symbol_protocols/symbol_toPrimitive_all_hints
// origin: languages/js/tests/js/test_symbol_protocols.rs

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
        if (hint === "number") return 42;
        if (hint === "string") return "forty-two";
        return true;
    }
};
__check(__line(+obj), "42");
__check(__line(`${obj}`), "forty-two");
__check(__line(obj + ""), "true");
__check(__line(obj > 40), "true");
