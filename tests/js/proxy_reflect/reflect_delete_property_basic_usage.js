// vybe-test: js/proxy_reflect/reflect_delete_property_basic_usage
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

const obj = { x: 1, y: 2 };
const ok = Reflect.deleteProperty(obj, "x");
__check(__line(ok), "true");
__check(__line(obj.x), "undefined");
