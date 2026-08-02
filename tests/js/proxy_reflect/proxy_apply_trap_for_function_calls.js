// vybe-test: js/proxy_reflect/proxy_apply_trap_for_function_calls
// origin: languages/js/tests/js/test_proxy_reflect.rs

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
        return target.apply(thisArg, args) * 2;
    }
};
function double(n) { return n * 2; }
const p = new Proxy(double, handler);
// apply trap doubles the already-doubled result
__check(__line(double(5)), "10");
__check(__line(typeof p), "function");
