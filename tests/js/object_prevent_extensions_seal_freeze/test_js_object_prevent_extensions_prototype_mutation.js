// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_prevent_extensions_prototype_mutation
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

const proto = { parentVal: 100 };
const obj = Object.create(proto);
Object.preventExtensions(obj);

try {
    // Modern ES6 Object.setPrototypeOf on non-extensible throws TypeError!
    Object.setPrototypeOf(obj, { newProto: 200 });
} catch (e) {
    __check(__line("SetPrototypeOf Failed"), "SetPrototypeOf Failed");
}
