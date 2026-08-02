// vybe-test: js/reactive_patterns/event_emitter_pattern
// origin: languages/js/tests/js/test_reactive_patterns.rs

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
    #listeners = new Map();
    on(event, fn) {
        if (!this.#listeners.has(event)) this.#listeners.set(event, []);
        this.#listeners.get(event).push(fn);
        return this;
    }
    off(event, fn) {
        const list = this.#listeners.get(event) || [];
        this.#listeners.set(event, list.filter(f => f !== fn));
        return this;
    }
    emit(event, ...args) {
        (this.#listeners.get(event) || []).forEach(fn => fn(...args));
        return this;
    }
    once(event, fn) {
        const wrapper = (...args) => { fn(...args); this.off(event, wrapper); };
        return this.on(event, wrapper);
    }
}
const ee = new EventEmitter();
const results = [];
ee.on("data", v => results.push("on:" + v));
ee.once("data", v => results.push("once:" + v));
ee.emit("data", 1);
ee.emit("data", 2);
console.log(results.join(","));
