// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_object_is_number_object_identity
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

const n1 = new Number(7);
const n2 = new Number(7);
__check(__line(`${n1 === n2}:${Object.is(n1, n2)}`), "false:false");
