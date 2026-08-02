// vybe-test: js/object_advanced/object_create_basic
// origin: languages/js/tests/js/test_object_advanced.rs

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

let proto = {
    greet() { return "Hello, " + this.name; }
};
let obj = Object.create(proto);
obj.name = "Alice";
__check(__line(obj.greet()), "Hello, Alice");
