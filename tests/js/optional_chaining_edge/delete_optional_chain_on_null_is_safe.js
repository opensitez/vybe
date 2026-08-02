// vybe-test: js/optional_chaining_edge/delete_optional_chain_on_null_is_safe
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

const obj = null;
// delete obj?.prop should not throw when obj is null
const result = delete obj?.prop;
__check(__line(result), "true");
