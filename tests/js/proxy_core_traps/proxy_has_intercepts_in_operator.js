// vybe-test: js/proxy_core_traps/proxy_has_intercepts_in_operator
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
    has(target, prop) {
        if (prop === "secret") return false;
        return prop in target;
    }
};
const obj = new Proxy({ secret: 42, public: 1 }, handler);
__check(__line("public" in obj), "true");
__check(__line("secret" in obj), "false");
