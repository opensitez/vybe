// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_arrow_function_this_unbound
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

const arrow = () => "arrow";
const proxy = new Proxy(arrow, {
    apply(target, thisArg, args) {
        return target.apply(thisArg, args);
    }
});
__check(__line(proxy.call({ custom: "ctx" })), "arrow");
