// vybe-test: js/object_descriptors/get_own_property_names_includes_non_enumerable
// origin: languages/js/tests/js/test_object_descriptors.rs

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
Object.defineProperty(obj, "hidden", { value: 1, enumerable: false, configurable: true });
obj.visible = 2;
const names = Object.getOwnPropertyNames(obj).sort();
__check(__line(names.join(",")), "hidden,visible");
