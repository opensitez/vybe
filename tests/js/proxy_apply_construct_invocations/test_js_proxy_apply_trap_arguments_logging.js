// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_apply_trap_arguments_logging
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

const log = [];
function add(a, b) { return a + b; }
const loggedAdd = new Proxy(add, {
    apply(target, thisArg, args) {
        log.push(args.join("+"));
        return target(...args);
    }
});
loggedAdd(1, 2);
loggedAdd(10, 20);
__check(__line(log.join(",")), "1+2,10+20");
