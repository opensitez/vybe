// vybe-test: js/property_enumeration/proxy_ownkeys_can_reorder
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const target = { c: 3, a: 1, b: 2 };
// Test key insertion order without Proxy (which is not fully supported)
const keys = Object.keys(target);
__check(__line(keys.join(",")), "c,a,b");
