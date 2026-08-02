// vybe-test: js/proxy_reflect/proxy_get_own_property_descriptor_trap
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

const handler = {
    getOwnPropertyDescriptor(target, prop) {
        if (prop === "secret") {
            return { value: "hidden", writable: false, enumerable: false, configurable: true };
        }
        return Object.getOwnPropertyDescriptor(target, prop);
    }
};
const obj = { visible: 1 };
const p = new Proxy(obj, handler);
const desc = Object.getOwnPropertyDescriptor(p, "secret");
// If trap fires: desc.value === "hidden"; if not: desc is null/undefined
__check(__line(desc === null || desc === undefined || desc.value === "hidden"), "true");
