// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_object_is_object_wrapper_distinct_instances
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

const n1 = new Number(5);
const n2 = new Number(5);
const b1 = new Boolean(false);
const b2 = new Boolean(false);
__check(__line(`${Object.is(n1, n2)}:${Object.is(b1, b2)}:${Object.is(n1, n1)}`), "false:false:true");
