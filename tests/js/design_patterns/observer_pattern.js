// vybe-test: js/design_patterns/observer_pattern
// origin: languages/js/tests/js/test_design_patterns.rs

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
    constructor() { this._events = {}; }
    on(event, listener) {
        (this._events[event] ??= []).push(listener);
        return this;
    }
    emit(event, ...args) {
        (this._events[event] ?? []).forEach(fn => fn(...args));
    }
    off(event, listener) {
        this._events[event] = (this._events[event] ?? []).filter(fn => fn !== listener);
    }
}
const emitter = new EventEmitter();
const log = [];
const handler = (x) => log.push(x);
emitter.on("data", handler);
emitter.emit("data", 1);
emitter.emit("data", 2);
emitter.off("data", handler);
emitter.emit("data", 3); // not received
console.log(log.join(","));
