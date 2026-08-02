// vybe-test: js/optional_chaining_edge/optional_chain_triggers_on_both_null_and_undefined
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

const a = null;
const b = undefined;
__check(__line(a?.x), "undefined");
__check(__line(b?.x), "undefined");
