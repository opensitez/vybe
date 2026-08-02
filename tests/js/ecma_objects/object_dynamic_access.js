// vybe-test: js/ecma_objects/object_dynamic_access
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

const obj = { foo: 1, bar: 2 };
const key = "foo";
__check(__line(obj[key]), "1");
obj["baz"] = 3;
__check(__line(obj.baz), "3");
