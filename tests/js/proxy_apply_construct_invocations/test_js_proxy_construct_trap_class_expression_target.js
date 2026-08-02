// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_class_expression_target
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

const ProxyAnonClass = new Proxy(class {
    constructor(val) { this.val = val; }
}, {
    construct(target, args) {
        return new target(args[0] * 100);
    }
});
const obj = new ProxyAnonClass(3);
__check(__line(obj.val), "300");
