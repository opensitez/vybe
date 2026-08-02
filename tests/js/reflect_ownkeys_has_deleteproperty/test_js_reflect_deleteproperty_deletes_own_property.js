// vybe-test: js/reflect_ownkeys_has_deleteproperty/test_js_reflect_deleteproperty_deletes_own_property
// origin: languages/js/tests/js/test_js_reflect_ownkeys_has_deleteproperty.rs

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

const obj = { a: 1, b: 2 };
const res = Reflect.deleteProperty(obj, "a");
__check(__line(res + "|hasA=" + ("a" in obj)), "true|hasA=false");
