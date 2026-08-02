// vybe-test: js/prototype_chain_advanced/delegation_pattern_via_object_create
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

const logger = {
    log(msg) { return `[LOG] ${msg}`; },
    error(msg) { return `[ERROR] ${msg}`; }
};
const app = Object.create(logger);
app.name = "MyApp";
app.run = function() { return this.log("Running " + this.name); };

__check(__line(app.run()), "[LOG] Running MyApp");
__check(__line(app.error("crash")), "[ERROR] crash");
