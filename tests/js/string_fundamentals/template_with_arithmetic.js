// vybe-test: js/string_fundamentals/template_with_arithmetic
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

const a = 3, b = 4;
__check(__line(`hypotenuse: ${Math.sqrt(a**2 + b**2)}`), "hypotenuse: 5");
