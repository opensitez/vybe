// vybe-test: js/object_prototype_methods/propertyisenumerable_symbol_enumerable
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

const s=Symbol("e"); const o={[s]:1}; __check(__line(o.propertyIsEnumerable(s)), "true");
