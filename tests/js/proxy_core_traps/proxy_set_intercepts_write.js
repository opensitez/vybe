// vybe-test: js/proxy_core_traps/proxy_set_intercepts_write
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

const log = [];
const handler = {
    set(target, prop, value) {
        log.push(`${prop}=${value}`);
        target[prop] = value;
        return true;
    }
};
const obj = new Proxy({}, handler);
obj.x = 1;
obj.y = 2;
__check(__line(log.join(",")), "x=1,y=2");
__check(__line(obj.x + obj.y), "3");
