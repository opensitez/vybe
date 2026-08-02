// vybe-test: js/class_private_advanced/static_public_class_field
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

class App {
    static name = "MyApp";
    static version = "2.0";
    static description() { return App.name + " v" + App.version; }
}
__check(__line(App.name), "MyApp");
__check(__line(App.version), "2.0");
__check(__line(App.description()), "MyApp v2.0");
