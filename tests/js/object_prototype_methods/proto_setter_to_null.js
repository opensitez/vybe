// vybe-test: js/object_prototype_methods/proto_setter_to_null
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

const o={x:1}; o.__proto__=null; __check(__line(o.x), "1"); __check(__line(Object.getPrototypeOf(o)), "null");
