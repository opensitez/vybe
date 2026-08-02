// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_bind_compatibility
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

function greet(prefix, suffix) { return `${prefix} ${this.name} ${suffix}`; }
const proxy = new Proxy(greet, {
    apply(target, thisArg, args) {
        return target.apply(thisArg, args);
    }
});
const bound = proxy.bind({ name: "World" }, "Hello");
__check(__line(bound("!")), "Hello World !");
