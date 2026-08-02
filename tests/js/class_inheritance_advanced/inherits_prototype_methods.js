// vybe-test: js/class_inheritance_advanced/inherits_prototype_methods
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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
    constructor() { this._handlers = {}; }
    on(event, fn) {
        (this._handlers[event] = this._handlers[event] || []).push(fn);
    }
    emit(event, ...args) {
        (this._handlers[event] || []).forEach(fn => fn(...args));
    }
}
class Button extends EventEmitter {
    click() { this.emit("click", this); }
}
const btn = new Button();
const log = [];
btn.on("click", () => log.push("clicked"));
btn.click();
btn.click();
console.log(log.join(","));
