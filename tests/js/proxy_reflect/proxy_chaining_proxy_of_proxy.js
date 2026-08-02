// vybe-test: js/proxy_reflect/proxy_chaining_proxy_of_proxy
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
const base = { value: 5 };
const inner = new Proxy(base, {
    get(target, prop) {
        log.push("inner:" + prop);
        return target[prop];
    }
});
const outer = new Proxy(inner, {
    get(target, prop) {
        log.push("outer:" + prop);
        return target[prop];
    }
});
const v = outer.value;
__check(__line(v), "5");
__check(__line(log[0]), "outer:value");
__check(__line(log[1]), "inner:value");
