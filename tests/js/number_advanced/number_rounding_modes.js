// vybe-test: js/number_advanced/number_rounding_modes
// origin: languages/js/tests/js/test_number_advanced.rs

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

// Round half up (standard)
const roundHalfUp = n => Math.floor(n + 0.5);
// Round half to even (banker's rounding simulation)
const roundHalfEven = n => {
    const floor = Math.floor(n);
    const frac = n - floor;
    if (Math.abs(frac - 0.5) < 1e-10) {
        return floor % 2 === 0 ? floor : floor + 1;
    }
    return Math.round(n);
};
console.log(roundHalfUp(2.5));
console.log(roundHalfUp(-2.5));
console.log(roundHalfEven(2.5));
console.log(roundHalfEven(3.5));
