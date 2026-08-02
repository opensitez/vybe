// vybe-test: js/destructuring_comprehensive/object_destructuring_nested_defaults
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

const config = { server: { host: "localhost" }, timeout: 3000 };
const { server: { host, port = 8080 }, timeout, retries = 3 } = config;
__check(__line(host), "localhost");
__check(__line(port), "8080");
__check(__line(timeout), "3000");
__check(__line(retries), "3");
