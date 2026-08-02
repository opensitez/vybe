// vybe-test: js/proxy_reflect/reflect_get_own_property_descriptor
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const obj = { x: 42 };
const desc = Reflect.getOwnPropertyDescriptor(obj, "x");
__check(__line(desc.value), "42");
__check(__line(desc.writable), "true");
__check(__line(desc.enumerable), "true");
