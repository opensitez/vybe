// vybe-test: js/optional_chaining_edge/optional_chain_does_not_trigger_on_zero_or_false
// origin: languages/js/tests/js/test_optional_chaining_edge.rs

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

const zero = 0;
const bool = false;
// 0?.x is different — 0 has no property 'x', but no short-circuit either
__check(__line(typeof zero?.toString), "function");
__check(__line(typeof bool?.toString), "function");
