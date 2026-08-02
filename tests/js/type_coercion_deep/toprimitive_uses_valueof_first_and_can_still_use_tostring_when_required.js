// vybe-test: js/type_coercion_deep/toprimitive_uses_valueof_first_and_can_still_use_tostring_when_required
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

const obj = {
    valueOf() {
        throw new Error("valueOf exploded");
    },
    toString() {
        return "stringified";
    }
};
try {
    console.log(+obj);
} catch (e) {
    console.log(e.message);
}
try {
    console.log(Number(obj));
} catch (e) {
    console.log(e.message);
}
console.log(String(obj));
