// vybe-test: js/class_private_deep/static_private_with_instance_methods
// origin: languages/js/tests/js/test_class_private_deep.rs

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

class EventLogger {
    static #log = [];
    static getLog() { return [...EventLogger.#log]; }
    log(event) { EventLogger.#log.push(event); }
}
const logger = new EventLogger();
logger.log("start");
logger.log("process");
logger.log("end");
__check(__line(EventLogger.getLog().join(",")), "start,process,end");
