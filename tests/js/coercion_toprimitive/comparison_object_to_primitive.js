// vybe-test: js/coercion_toprimitive/comparison_object_to_primitive
// origin: languages/js/tests/js/test_coercion_toprimitive.rs

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

// Date uses valueOf (timestamp) for comparisons
const d1 = new Date(0);
const d2 = new Date(1000);
console.log(d1 < d2);
console.log(d2 - d1);
