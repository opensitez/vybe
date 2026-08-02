// vybe-test: js/prototype_chain_shadowing_property_lookup/test_js_prototype_chain_setter_on_prototype_intercepts_assignment
// origin: languages/js/tests/js/test_js_prototype_chain_shadowing_property_lookup.rs

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

const proto = {
    set count(v) { this._count = v * 10; },
    get count() { return this._count; }
};
const obj = Object.create(proto);
obj.count = 5; // Triggers prototype setter with 'this' pointing to obj!
__check(__line(obj._count + "|" + (Object.hasOwn(obj, "count"))), "50|false");
