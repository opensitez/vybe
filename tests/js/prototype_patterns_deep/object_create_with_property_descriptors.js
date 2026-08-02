// vybe-test: js/prototype_patterns_deep/object_create_with_property_descriptors
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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

const proto = { greet() { return "Hello from " + this.name; } };
const obj = Object.create(proto, {
    name: { value: "Alice", writable: true, enumerable: true, configurable: true }
});
__check(__line(obj.name), "Alice");
__check(__line(obj.greet()), "Hello from Alice");
