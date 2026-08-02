// vybe-test: js/object_advanced/method_shorthand_in_object
// origin: languages/js/tests/js/test_object_advanced.rs

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

let calc = {
    add(a, b) { return a + b; },
    mul(a, b) { return a * b; }
};
__check(__line(calc.add(3, 4)), "7");
__check(__line(calc.mul(3, 4)), "12");
