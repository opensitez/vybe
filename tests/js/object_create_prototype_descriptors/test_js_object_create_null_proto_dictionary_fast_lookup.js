// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_null_proto_dictionary_fast_lookup
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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

const dict = Object.create(null);
dict.key1 = "v1";
dict["__proto__"] = "not_a_proto";
__check(__line(dict["__proto__"] + "|" + (Object.getPrototypeOf(dict) === null)), "not_a_proto|true");
