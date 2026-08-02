// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_same_value_zero_array_index_of_vs_includes
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

const arr = [NaN];
__check(__line(`${arr.indexOf(NaN)}:${arr.includes(NaN)}`), "-1:true"); // indexOf uses === (returns -1), includes uses SameValueZero (returns true)!
