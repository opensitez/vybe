// vybe-test: js/proxy_reflect/proxy_has_trap_intercepts_in_operator
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

const range = { min: 1, max: 10 };
const handler = {
    has(target, prop) {
        const n = Number(prop);
        return n >= target.min && n <= target.max;
    }
};
const p = new Proxy(range, handler);
__check(__line(5 in p), "true");
__check(__line(15 in p), "false");
