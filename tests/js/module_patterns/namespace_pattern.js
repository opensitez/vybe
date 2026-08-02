// vybe-test: js/module_patterns/namespace_pattern
// origin: languages/js/tests/js/test_module_patterns.rs

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

const App = {
    utils: {
        add: (a, b) => a + b,
        multiply: (a, b) => a * b,
    },
    config: {
        version: "1.0",
        debug: false,
    },
    init() {
        return `App v${this.config.version} initialized`;
    }
};
__check(__line(App.utils.add(3, 4)), "7");
__check(__line(App.init()), "App v1.0 initialized");
__check(__line(App.config.debug), "false");
