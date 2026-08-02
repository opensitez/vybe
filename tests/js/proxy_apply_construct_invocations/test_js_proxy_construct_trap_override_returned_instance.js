// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_override_returned_instance
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

function Person(name) { this.name = name; }
const proxy = new Proxy(Person, {
    construct(target, args) {
        return { custom: true, name: args[0].toUpperCase() };
    }
});
const p = new proxy("charlie");
__check(__line(p.name + "|" + p.custom), "CHARLIE|true");
