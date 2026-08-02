// vybe-test: js/inheritance/test_25_default_param_values
// origin: languages/js/tests/js/js_inheritance_test.rs

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

class Config {
            constructor(host = "localhost", port = 8080) {
                this.host = host;
                this.port = port;
            }
        }
        let c1 = new Config();
        let c2 = new Config("example.com", 3000);
        __check(__line(c1.host, c1.port), "localhost 8080");
        __check(__line(c2.host, c2.port), "example.com 3000");
