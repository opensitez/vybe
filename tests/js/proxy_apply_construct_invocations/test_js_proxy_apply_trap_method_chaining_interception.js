// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_method_chaining_interception
// origin: languages/js/tests/js/test_js_proxy_apply_construct_invocations.rs

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

const obj = {
    val: 0,
    add(n) { this.val += n; return this; }
};
const proxy = new Proxy(obj, {
    get(target, prop, receiver) {
        const orig = target[prop];
        if (typeof orig === "function") {
            return new Proxy(orig, {
                apply(fnTarget, thisArg, args) {
                    return fnTarget.apply(target, args);
                }
            });
        }
        return orig;
    }
});
proxy.add(5).add(10);
__check(__line(obj.val), "15");
