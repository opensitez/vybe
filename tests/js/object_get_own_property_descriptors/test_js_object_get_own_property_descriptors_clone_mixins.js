// vybe-test: js/object_get_own_property_descriptors/test_js_object_get_own_property_descriptors_clone_mixins
// origin: languages/js/tests/js/test_js_object_get_own_property_descriptors.rs

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

const source = {
    _count: 0,
    get count() { return this._count; },
    set count(v) { this._count = v; }
};
const clone = Object.create(Object.getPrototypeOf(source), Object.getOwnPropertyDescriptors(source));
clone.count = 50;
__check(__line(clone.count + "|" + source.count), "50|0");
