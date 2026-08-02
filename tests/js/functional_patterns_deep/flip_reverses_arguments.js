// vybe-test: js/functional_patterns_deep/flip_reverses_arguments
// origin: languages/js/tests/js/test_functional_patterns_deep.rs

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

const flip = fn => (a, b, ...rest) => fn(b, a, ...rest);
const subtract = (a, b) => a - b;
const flipped = flip(subtract);
__check(__line(subtract(10, 3)), "7");
__check(__line(flipped(10, 3)), "-7");
