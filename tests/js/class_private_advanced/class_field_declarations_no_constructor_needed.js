// vybe-test: js/class_private_advanced/class_field_declarations_no_constructor_needed
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    host = "localhost";
    port = 8080;
    secure = false;
    toString() {
        const proto = this.secure ? "https" : "http";
        return proto + "://" + this.host + ":" + this.port;
    }
}
const cfg = new Config();
__check(__line(cfg.toString()), "http://localhost:8080");
cfg.port = 443;
cfg.secure = true;
__check(__line(cfg.toString()), "https://localhost:443");
