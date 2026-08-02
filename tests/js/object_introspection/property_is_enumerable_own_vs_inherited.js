// vybe-test: js/object_introspection/property_is_enumerable_own_vs_inherited
// origin: languages/js/tests/js/test_object_introspection.rs

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

const proto = { fromProto: 1 };
const obj = Object.create(proto);
obj.ownEnum = 2;
Object.defineProperty(obj, "ownHidden", { value: 3, enumerable: false });
__check(__line(obj.propertyIsEnumerable("ownEnum")), "true");
__check(__line(obj.propertyIsEnumerable("ownHidden")), "false");
__check(__line(obj.propertyIsEnumerable("fromProto")), "false");
