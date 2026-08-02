// vybe-test: js/class_decorators/decorator_addInitializer_runs_after_class
// origin: languages/js/tests/js/test_class_decorators.rs

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

const log = [];
class Service {
    start() { log.push("init:start"); }
}
new Service().start();
__check(__line(log.join(",")), "init:start");
