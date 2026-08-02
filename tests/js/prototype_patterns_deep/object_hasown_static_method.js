// vybe-test: js/prototype_patterns_deep/object_hasown_static_method
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

const obj = { a: 1 };
const nullProto = Object.create(null);
nullProto.key = "value";
__check(__line(Object.hasOwn(obj, "a")), "true");
// Object.hasOwn works even on null-prototype objects
__check(__line(Object.hasOwn(nullProto, "key")), "true");
