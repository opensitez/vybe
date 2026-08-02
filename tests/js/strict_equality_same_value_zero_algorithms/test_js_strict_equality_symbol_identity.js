// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_strict_equality_symbol_identity
// origin: languages/js/tests/js/test_js_strict_equality_same_value_zero_algorithms.rs

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

const s1 = Symbol("key");
const s2 = Symbol("key");
__check(__line(`${s1 === s1}:${s1 === s2}`), "true:false");
