// vybe-test: js/misc_advanced_patterns/proxy_logging_trap
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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

function logged(obj) {
    const log = [];
    const proxy = new Proxy(obj, {
        get(target, prop, receiver) {
            log.push("get:" + String(prop));
            return Reflect.get(target, prop, receiver);
        },
        set(target, prop, value, receiver) {
            log.push("set:" + String(prop));
            return Reflect.set(target, prop, value, receiver);
        }
    });
    return [proxy, log];
}
const [p, log] = logged({ x: 1 });
p.x;
p.y = 2;
p.x;
__check(__line(log.join(",")), "get:x,set:y,get:x");
