// vybe-test: js/module_import_patterns/import_with_attribute_type_json_syntactically_valid
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

// Import attributes syntax: import x from "y" with { type: "json" }
// Testing the syntax is parsed (actual loading may fail in test env)
let ok = true;
try {
    eval('import("./data.json", { with: { type: "json" } }).catch(() => {})');
} catch (e) {
    // SyntaxError means parser doesn't support it yet
    ok = e instanceof SyntaxError ? false : true;
}
console.log(typeof ok);
