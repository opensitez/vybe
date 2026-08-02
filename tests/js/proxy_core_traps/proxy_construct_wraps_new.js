// vybe-test: js/proxy_core_traps/proxy_construct_wraps_new
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

class Point { constructor(x, y) { this.x = x; this.y = y; } }
const handler = {
    construct(target, args) {
        const instance = new target(...args);
        instance.created = true;
        return instance;
    }
};
const ProxiedPoint = new Proxy(Point, handler);
const p = new ProxiedPoint(1, 2);
__check(__line(p.x), "1");
__check(__line(p.created), "true");
