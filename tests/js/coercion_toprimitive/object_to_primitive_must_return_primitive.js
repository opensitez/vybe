// vybe-test: js/coercion_toprimitive/object_to_primitive_must_return_primitive
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

const bad = {
    valueOf() {
        throw new Error("valueOf boom");
    },
    toString() {
        return "ok";
    }
};
try {
    console.log(+bad);
} catch (e) {
    console.log(e.message);
}
console.log(String(bad));
