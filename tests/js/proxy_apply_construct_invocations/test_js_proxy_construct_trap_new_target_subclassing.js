// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_new_target_subclassing
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

class Base {
    constructor() { this.base = true; }
}
let capturedNewTarget;
const proxy = new Proxy(Base, {
    construct(target, args, newTarget) {
        capturedNewTarget = newTarget;
        return Reflect.construct(target, args, newTarget);
    }
});
class Derived extends proxy {}
const instance = new Derived();
__check(__line(capturedNewTarget === Derived), "true");
