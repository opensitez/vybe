// vybe-test: js/function_prototype_metadata/getter_setter_function_name_prefixes
// origin: languages/js/tests/js/test_function_prototype_metadata.rs

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

const obj = { get prop() {}, set prop(v) {} }; const desc = Object.getOwnPropertyDescriptor(obj, "prop"); __check(__line(desc.get.name + "|" + desc.set.name), "get prop|set prop");
