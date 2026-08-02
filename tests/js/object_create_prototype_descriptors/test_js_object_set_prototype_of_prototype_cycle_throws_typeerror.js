// vybe-test: js/object_create_prototype_descriptors/test_js_object_set_prototype_of_prototype_cycle_throws_typeerror
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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

const a = {};
const b = Object.create(a);
try {
    Object.setPrototypeOf(a, b); // Creating a prototype cycle is a TypeError!
} catch (e) {
    __check(__line("Prototype Cycle TypeError"), "Prototype Cycle TypeError");
}
