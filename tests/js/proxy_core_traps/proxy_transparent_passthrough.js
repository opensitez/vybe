// vybe-test: js/proxy_core_traps/proxy_transparent_passthrough
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

const target = { x: 1, y: 2 };
const proxy = new Proxy(target, {});
proxy.z = 3;
__check(__line(proxy.x), "1");
__check(__line(proxy.z), "3");
__check(__line(target.z), "3"); // writes go to target
