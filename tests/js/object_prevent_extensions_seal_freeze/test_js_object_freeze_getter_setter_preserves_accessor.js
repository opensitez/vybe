// vybe-test: js/object_prevent_extensions_seal_freeze/test_js_object_freeze_getter_setter_preserves_accessor
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

const obj = {
    _count: 0,
    get count() { return this._count; },
    set count(v) { this._count = v; }
};
Object.freeze(obj);
// Calling setter mutates backing field because accessor properties are frozen without changing getter/setter pointers
obj.count = 10;
__check(__line(obj.count), "10");
