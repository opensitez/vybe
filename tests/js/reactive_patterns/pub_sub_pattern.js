// vybe-test: js/reactive_patterns/pub_sub_pattern
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

class PubSub {
    #topics = new Map();
    subscribe(topic, fn) {
        if (!this.#topics.has(topic)) this.#topics.set(topic, new Set());
        this.#topics.get(topic).add(fn);
        return () => this.#topics.get(topic).delete(fn);
    }
    publish(topic, data) {
        (this.#topics.get(topic) || new Set()).forEach(fn => fn(data));
    }
}
const ps = new PubSub();
const log = [];
const unsub = ps.subscribe("news", msg => log.push(msg));
ps.subscribe("news", msg => log.push("B:" + msg));
ps.publish("news", "hello");
unsub();
ps.publish("news", "world");
console.log(log.join(","));
