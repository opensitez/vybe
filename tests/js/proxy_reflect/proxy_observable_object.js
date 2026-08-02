// vybe-test: js/proxy_reflect/proxy_observable_object
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

const changes = [];
const state = new Proxy({ count: 0 }, {
    set(obj, prop, value) {
        const old = obj[prop];
        obj[prop] = value;
        changes.push(prop + ":" + old + "->" + value);
        return true;
    }
});
state.count = 1;
state.count = 2;
__check(__line(changes[0]), "count:0->1");
__check(__line(changes[1]), "count:1->2");
