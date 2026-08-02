// vybe-test: js/proxy_core_traps/proxy_read_only_property
// origin: languages/js/tests/js/test_proxy_core_traps.rs

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
    set(target, prop, value) {
        if (prop === "immutable") {
            throw new TypeError("Cannot set immutable");
        }
        target[prop] = value;
        return true;
    }
};
const obj = new Proxy({ immutable: 42, mutable: 0 }, handler);
let threw = false;
try { obj.immutable = 99; } catch { threw = true; }
__check(__line(threw), "true");
obj.mutable = 10;
__check(__line(obj.mutable), "10");
