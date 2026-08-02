// vybe-test: js/coercion_toprimitive/boolean_coercion_falsy_values
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

const falsies = [false, 0, "", null, undefined, NaN, -0, 0n];
__check(__line(falsies.every(v => !v)), "true");
const truthy = [1, "a", {}, [], () => {}, Infinity];
__check(__line(truthy.every(v => !!v)), "true");
