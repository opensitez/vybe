// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_strict_equality_boxed_primitives
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

__check(__line(`${new Number(7) === 7}:${new Number(7) === new Number(7)}:${Object.is(new Number(7), new Number(7))}`), "false:false:false");
__check(__line(`${new Boolean(true) === true}:${new String("x") === "x"}`), "false:false");
