// vybe-test: js/strict_equality_same_value_zero_algorithms/test_js_object_identity_distinguishes_boxed_wrappers
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

const set = new Set();
const a = new Number(7);
const b = new Number(7);
set.add(a);
set.add(b);

const map = new Map();
const o1 = { x: 1 };
const o2 = { x: 1 };
map.set(o1, "first");
map.set(o2, "second");

__check(__line(`${set.size}:${map.size}`), "2:2");
