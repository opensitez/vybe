// vybe-test: js/coercion_toprimitive/valueof_tostring_are_tied_to_primitive_hints
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

const trace = [];
const obj = {
    valueOf() {
        trace.push("valueOf");
        return 4;
    },
    toString() {
        trace.push("toString");
        return "9";
    }
};
__check(__line(+obj), "4");
__check(__line(`${obj}`), "9");
__check(__line(trace.join(",")), "valueOf,toString");
