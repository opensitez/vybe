// vybe-test: js/function_prototype_deep/call_with_symbol_property_on_this
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

const key = Symbol("k"); const obj = { [key]: 7 }; function read() { return this[key]; } __check(__line(read.call(obj)), "7");
