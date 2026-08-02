// vybe-test: js/oop_patterns_advanced/event_sourcing_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

class EventStore {
    #events = [];
    append(event) { this.#events.push({...event, timestamp: Date.now()}); }
    replay(handlers) {
        let state = {};
        for (const event of this.#events) {
            const handler = handlers[event.type];
            if (handler) state = handler(state, event);
        }
        return state;
    }
}
const store = new EventStore();
store.append({ type: "Created", name: "Alice" });
store.append({ type: "Updated", field: "age", value: 30 });
const state = store.replay({
    Created: (s, e) => ({ ...s, name: e.name }),
    Updated: (s, e) => ({ ...s, [e.field]: e.value }),
});
console.log(state.name);
console.log(state.age);
