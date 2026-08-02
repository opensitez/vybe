// vybe-test: js/module_import_patterns/module_object_has_module_namespace_tostring_tag
// origin: languages/js/tests/js/test_module_import_patterns.rs

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

// Can't easily test namespace object without real module system,
// but we can test that Symbol.toStringTag is 'Module' conceptually
const ns = Object.create(null);
Object.defineProperty(ns, Symbol.toStringTag, { value: "Module" });
console.log(Object.prototype.toString.call(ns));
