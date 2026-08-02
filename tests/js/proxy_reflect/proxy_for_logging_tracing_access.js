// vybe-test: js/proxy_reflect/proxy_for_logging_tracing_access
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

const log = [];
const handler = {
    get(target, prop) {
        log.push("get:" + prop);
        return target[prop];
    },
    set(target, prop, value) {
        log.push("set:" + prop);
        target[prop] = value;
        return true;
    }
};
const obj = new Proxy({}, handler);
obj.name = "Alice";
const n = obj.name;
__check(__line(log[0]), "set:name");
__check(__line(log[1]), "get:name");
__check(__line(n), "Alice");
