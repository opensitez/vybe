// vybe-test: js/module_import_patterns/destructure_namespace_import_simulation
// origin: languages/js/tests/js/test_module_import_patterns.rs

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

// Simulate named exports with object
const mathModule = {
    add: (a, b) => a + b,
    sub: (a, b) => a - b,
    PI: 3.14159
};
const { add, sub, PI } = mathModule;
__check(__line(add(2, 3)), "5");
__check(__line(sub(5, 2)), "3");
__check(__line(PI.toFixed(2)), "3.14");
