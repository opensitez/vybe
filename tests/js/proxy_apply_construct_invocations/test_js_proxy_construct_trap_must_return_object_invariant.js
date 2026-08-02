// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_must_return_object_invariant
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

function Item() {}
const proxy = new Proxy(Item, {
    construct(target, args) {
        return 42; // Construct trap MUST return an object!
    }
});
try {
    new proxy();
} catch (e) {
    __check(__line("Construct Non-Object Invariant Error"), "Construct Non-Object Invariant Error");
}
