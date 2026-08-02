// vybe-test: js/ecma_operators/optional_chaining_computed_property
// origin: languages/js/tests/js/test_ecma_operators.rs

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

const obj = { nested: { value: 7 } };
const key = "nested";
__check(__line(obj?.[key]?.value), "7");
__check(__line(obj?.["missing"]?.value), "undefined");
