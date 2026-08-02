// vybe-test: js/proxy_traps_advanced/proxy_has_trap_can_hide_properties
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

const hidden = new Set(["secret"]);
const proxy = new Proxy(
    { secret: 42, public: 1 },
    { has(target, prop) { return !hidden.has(prop) && prop in target; } }
);
__check(__line("public" in proxy), "true");
__check(__line("secret" in proxy), "false");
