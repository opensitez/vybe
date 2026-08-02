// vybe-test: js/proxy_getownpropertydescriptor_defineproperty/test_js_proxy_defineproperty_trap_symbol_property
// origin: languages/js/tests/js/test_js_proxy_getownpropertydescriptor_defineproperty.rs

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

const sym = Symbol("id");
const target = {};
const proxy = new Proxy(target, {
    defineProperty(t, prop, desc) {
        return Reflect.defineProperty(t, prop, desc);
    }
});
Object.defineProperty(proxy, sym, { value: "SymbolDefined" });
__check(__line(target[sym]), "SymbolDefined");
