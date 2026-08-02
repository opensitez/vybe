// vybe-test: js/module_import_patterns/lazy_initialization_with_import_simulation
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

// Simulate lazy singleton loading
let instance = null;
function getInstance() {
    if (!instance) instance = { value: Math.random() };
    return instance;
}
const a = getInstance();
const b = getInstance();
__check(__line(a === b), "true");
