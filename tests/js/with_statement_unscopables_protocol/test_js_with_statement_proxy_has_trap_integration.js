// vybe-test: js/with_statement_unscopables_protocol/test_js_with_statement_proxy_has_trap_integration
// origin: languages/js/tests/js/test_js_with_statement_unscopables_protocol.rs

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

const target = { a: 1 };
const proxy = new Proxy(target, {
    has(t, prop) {
        if (prop === "b") return true;
        return Reflect.has(t, prop);
    },
    get(t, prop) {
        if (prop === "b") return "TrappedB";
        return Reflect.get(t, prop);
    }
});
with (proxy) {
    __check(__line(a + "|" + b), "1|TrappedB");
}
