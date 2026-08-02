// vybe-test: js/misc_advanced_patterns/object_create_descriptors
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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

const proto = {
    greet() { return `Hello, ${this.name}`; }
};
const obj = Object.create(proto, {
    name: { value: "Alice", writable: true, enumerable: true, configurable: true },
    age: { value: 30, writable: false, enumerable: true, configurable: false }
});
__check(__line(obj.greet()), "Hello, Alice");
obj.name = "Bob";
__check(__line(obj.name), "Bob");
obj.age = 99;
__check(__line(obj.age), "30");
