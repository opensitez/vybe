// vybe-test: js/object_prototype_methods/hasownproperty_numeric_string_key
// origin: languages/js/tests/js/test_object_prototype_methods.rs

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

const o={"0":1}; __check(__line(o.hasOwnProperty(0)), "true"); __check(__line(o.hasOwnProperty("0")), "true");
