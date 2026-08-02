// vybe-test: js/class_patterns/observer_pattern
// origin: languages/js/tests/js/test_class_patterns.rs

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

class EventEmitter {
    constructor() { this.listeners = {}; }
    on(event, fn) {
        if (!this.listeners[event]) this.listeners[event] = [];
        this.listeners[event].push(fn);
    }
    emit(event, ...args) {
        if (this.listeners[event]) {
            this.listeners[event].forEach(fn => fn(...args));
        }
    }
}
let emitter = new EventEmitter();
emitter.on("data", val => console.log("got: " + val));
emitter.on("data", val => console.log("also: " + val));
emitter.emit("data", 42);
