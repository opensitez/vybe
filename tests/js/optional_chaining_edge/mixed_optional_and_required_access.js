// vybe-test: js/optional_chaining_edge/mixed_optional_and_required_access
// origin: languages/js/tests/js/test_optional_chaining_edge.rs

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

const config = {
    server: { host: "localhost", port: 8080 }
};
// Only the first access is optional
__check(__line(config?.server.host), "localhost");
__check(__line(config?.missing?.port), "undefined");
