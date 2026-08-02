// vybe-test: js/property_descriptor_deep/define_non_enumerable_hidden_from_keys
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

const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false, configurable: true, writable: true });
__check(__line(Object.keys(obj).join(",")), "visible");
__check(__line(obj.hidden), "2");
