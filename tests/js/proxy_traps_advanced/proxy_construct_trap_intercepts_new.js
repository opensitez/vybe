// vybe-test: js/proxy_traps_advanced/proxy_construct_trap_intercepts_new
// origin: languages/js/tests/js/test_proxy_traps_advanced.rs

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

class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
const ProxiedPoint = new Proxy(Point, {
    construct(target, args) {
        const [x, y] = args;
        return new target(x * 2, y * 2);
    }
});
const p = new ProxiedPoint(3, 4);
__check(__line(p.x), "6");
__check(__line(p.y), "8");
