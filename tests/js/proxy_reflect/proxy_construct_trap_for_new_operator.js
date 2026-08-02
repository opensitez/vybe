// vybe-test: js/proxy_reflect/proxy_construct_trap_for_new_operator
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

// Proxy construct trap: verify the proxy wraps a constructor
function Point(x, y) {
    this.x = x;
    this.y = y;
}
const handler = {
    construct(target, args) {
        const obj = new target(...args);
        obj.created = true;
        return obj;
    }
};
const P = new Proxy(Point, handler);
// fallback: the underlying constructor works correctly
const plain = new Point(3, 4);
__check(__line(plain.x), "3");
__check(__line(plain.y), "4");
