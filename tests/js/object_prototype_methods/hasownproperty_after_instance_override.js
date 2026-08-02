// vybe-test: js/object_prototype_methods/hasownproperty_after_instance_override
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

const o={hasOwnProperty(){return false}}; __check(__line(o.hasOwnProperty("x")), "false"); __check(__line(Object.prototype.hasOwnProperty.call(o,"hasOwnProperty")), "true");
