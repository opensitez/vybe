// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_prevent_extensions_array_push_throws
// origin: languages/js/tests/js/test_js_object_prevent_extensions_seal_freeze.rs

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

const arr = [1, 2];
Object.preventExtensions(arr);
try {
    arr.push(3);
} catch (e) {
    __check(__line("Push Non-Extensible Array Error"), "Push Non-Extensible Array Error");
}
