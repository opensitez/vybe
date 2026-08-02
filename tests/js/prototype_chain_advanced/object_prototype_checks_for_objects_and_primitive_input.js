// vybe-test: js/prototype_chain_advanced/object_prototype_checks_for_objects_and_primitive_input
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

const proto = {};
const obj = Object.create(proto);
console.log(Object.getPrototypeOf(proto) === Object.prototype);
console.log(Object.prototype.isPrototypeOf(obj));
try {
    console.log(Object.prototype.isPrototypeOf(1));
} catch (e) {
    console.log("PrimitiveTypeError");
}
