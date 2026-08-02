// vybe-test: js/proxy_reflect/proxy_as_default_values_provider
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

function withDefaults(target, defaults) {
    return new Proxy(target, {
        get(t, prop) {
            return prop in t ? t[prop] : defaults[prop];
        }
    });
}
const config = withDefaults({ port: 8080 }, { host: "localhost", port: 3000, debug: false });
__check(__line(config.port), "8080");
__check(__line(config.host), "localhost");
__check(__line(config.debug), "false");
