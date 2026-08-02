// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_strict_equality_object_references
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

const o1 = { x: 1 };
const o2 = { x: 1 };
const o3 = o1;
__check(__line(`${o1 === o2}:${o1 === o3}`), "false:true");
