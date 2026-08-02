// vybe-test: js/proxy_traps_advanced/proxy_apply_trap_wraps_function_call
// origin: languages/js/tests/js/test_proxy_traps_advanced.rs

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

function add(a, b) { return a + b; }
const proxy = new Proxy(add, {
    apply(target, thisArg, args) {
        __check(__line("called with " + args.length + " args"), "called with 2 args");
        return target.apply(thisArg, args);
    }
});
__check(__line(proxy(2, 3)), "5");
