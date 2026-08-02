// vybe-test: js/object_advanced/computed_keys_dynamic
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

let field = "name";
let obj = { [field]: "Alice", [field + "Length"]: 5 };
__check(__line(obj.name), "Alice");
__check(__line(obj.nameLength), "5");
