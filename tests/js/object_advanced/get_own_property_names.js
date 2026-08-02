// vybe-test: js/object_advanced/get_own_property_names
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

let obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
__check(__line(Object.keys(obj).join(",")), "a,b");
__check(__line(Object.getOwnPropertyNames(obj).join(",")), "a,b,hidden");
