// vybe-test: js/reactive_patterns/reactive_store
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

class Store {
    #state;
    #subscribers = [];
    constructor(init) { this.#state = init; }
    getState() { return this.#state; }
    setState(updater) {
        this.#state = typeof updater === "function" ? updater(this.#state) : { ...this.#state, ...updater };
        this.#subscribers.forEach(fn => fn(this.#state));
    }
    subscribe(fn) {
        this.#subscribers.push(fn);
        return () => { this.#subscribers = this.#subscribers.filter(s => s !== fn); };
    }
}
const store = new Store({ count: 0 });
const log = [];
const unsub = store.subscribe(s => log.push(s.count));
store.setState(s => ({ count: s.count + 1 }));
store.setState(s => ({ count: s.count + 1 }));
unsub();
store.setState(s => ({ count: s.count + 1 }));
console.log(log.join(","));
console.log(store.getState().count);
