// vybe-test: js/object_advanced/define_property_getter_setter
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

let obj = { _name: "Alice" };
Object.defineProperty(obj, "name", {
    get() { return this._name.toUpperCase(); },
    set(val) { this._name = val; }
});
__check(__line(obj.name), "ALICE");
obj.name = "Bob";
__check(__line(obj.name), "BOB");
