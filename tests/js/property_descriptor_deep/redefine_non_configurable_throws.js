// vybe-test: js/property_descriptor_deep/redefine_non_configurable_throws
// origin: languages/js/tests/js/test_property_descriptor_deep.rs

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

const obj = {};
Object.defineProperty(obj, "p", { value: 1, configurable: false });
let threw = false;
try {
    Object.defineProperty(obj, "p", { value: 2 });
} catch {
    threw = true;
}
__check(__line(threw), "true");
