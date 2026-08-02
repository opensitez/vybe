// vybe-test: js/module_patterns/dependency_injection_pattern
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

class Logger {
    log(msg) { return "[LOG] " + msg; }
}
class UserService {
    constructor(logger) { this.logger = logger; }
    createUser(name) {
        const msg = this.logger.log(`Creating user: ${name}`);
        return msg;
    }
}
const logger = new Logger();
const service = new UserService(logger);
__check(__line(service.createUser("Alice")), "[LOG] Creating user: Alice");
