// vybe-test: js/module_import_patterns/dynamic_import_returns_promise
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

// We can check that import() returns a Promise (even if module doesn't exist,
// the call site returns a thenable)
const maybePromise = import("./nonexistent_module_abc.js").catch(() => "failed");
console.log(maybePromise instanceof Promise);
