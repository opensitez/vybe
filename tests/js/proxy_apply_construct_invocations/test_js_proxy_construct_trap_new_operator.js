// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_new_operator
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

function User(name) { this.name = name; }
const proxy = new Proxy(User, {
    construct(target, args, newTarget) {
        const obj = new target(...args);
        obj.created = true;
        return obj;
    }
});
const u = new proxy("Bob");
__check(__line(u.name + "|" + u.created), "Bob|true");
