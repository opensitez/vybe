// vybe-test: js/object_methods_deep/object_create_with_prototype
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const proto = { greet() { return "hello from " + this.name; } };
const obj = Object.create(proto);
obj.name = "World";
__check(__line(obj.greet()), "hello from World");
__check(__line(Object.getPrototypeOf(obj) === proto), "true");
