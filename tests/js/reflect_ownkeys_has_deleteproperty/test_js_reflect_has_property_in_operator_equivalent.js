// vybe-test: js/reflect_ownkeys_has_deleteproperty/test_js_reflect_has_property_in_operator_equivalent
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

const proto = { parentKey: 10 };
const obj = Object.create(proto);
obj.ownKey = 20;
__check(__line(Reflect.has(obj, "ownKey") + "|" + Reflect.has(obj, "parentKey") + "|" + Reflect.has(obj, "missing")), "true|true|false");
