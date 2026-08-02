// vybe-test: js/ecma_objects/object_dynamic_property_delete_and_readd
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const obj = { a: 1, b: 2 };
delete obj["a"];
obj["a"] = 3;
__check(__line(obj.a), "3");
__check(__line(Object.keys(obj).join(",")), "b,a");
