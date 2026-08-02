// vybe-test: js/proxy_core_traps/proxy_apply_wraps_function
// origin: languages/js/tests/js/test_proxy_core_traps.rs

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

const handler = {
    apply(target, thisArg, args) {
        return target(...args) * 2;
    }
};
const double = new Proxy((x) => x + 1, handler);
__check(__line(double(5)), "12"); // (5+1)*2 = 12
