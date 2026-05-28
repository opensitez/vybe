/// Reactive patterns — observable, pub-sub, streams

use super::helpers::run_js;

#[test]
fn simple_observable() {
    assert_eq!(run_js(r#"
class Observable {
    constructor(subscribe) { this._subscribe = subscribe; }
    subscribe(observer) { return this._subscribe(observer); }
    map(fn) {
        return new Observable(obs => this.subscribe({
            next: v => obs.next(fn(v)),
            error: e => obs.error(e),
            complete: () => obs.complete()
        }));
    }
    filter(fn) {
        return new Observable(obs => this.subscribe({
            next: v => fn(v) && obs.next(v),
            error: e => obs.error(e),
            complete: () => obs.complete()
        }));
    }
    static of(...values) {
        return new Observable(obs => {
            values.forEach(v => obs.next(v));
            obs.complete();
        });
    }
}
const results = [];
Observable.of(1,2,3,4,5)
    .filter(x => x % 2 === 0)
    .map(x => x * 10)
    .subscribe({ next: v => results.push(v), error: ()=>{}, complete: ()=>{} });
console.log(results.join(","));
"#), vec!["20,40"]);
}

#[test]
fn event_emitter_pattern() {
    assert_eq!(run_js(r#"
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
"#), vec!["on:1,once:1,on:2"]);
}

#[test]
fn reactive_store() {
    assert_eq!(run_js(r#"
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
"#), vec!["1,2", "3"]);
}

#[test]
fn pub_sub_pattern() {
    assert_eq!(run_js(r#"
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
"#), vec!["hello,B:hello,B:world"]);
}

#[test]
fn signal_reactivity() {
    assert_eq!(run_js(r#"
function createSignal(init) {
    let value = init;
    const subscribers = new Set();
    const get = () => value;
    const set = (v) => { value = v; subscribers.forEach(fn => fn(v)); };
    const subscribe = fn => { subscribers.add(fn); return () => subscribers.delete(fn); };
    return [get, set, subscribe];
}
function computed(deps, fn) {
    const [get, set] = createSignal(fn(...deps.map(d => d())));
    deps.forEach(dep => dep[2](() => set(fn(...deps.map(d => d())))));
    return get;
}
const [count, setCount, subCount] = createSignal(0);
const doubled = computed([[count, null, subCount]], c => c * 2);
const log = [];
subCount(v => log.push("count:" + v));
setCount(5);
setCount(10);
console.log(log.join(","));
console.log(doubled());
"#), vec!["count:5,count:10", "20"]);
}

#[test]
fn stream_accumulator() {
    assert_eq!(run_js(r#"
class Stream {
    #buffer = [];
    #subscribers = [];
    push(...values) {
        this.#buffer.push(...values);
        values.forEach(v => this.#subscribers.forEach(fn => fn(v)));
        return this;
    }
    subscribe(fn) { this.#subscribers.push(fn); return this; }
    collect() { return [...this.#buffer]; }
    map(fn) {
        const out = new Stream();
        this.subscribe(v => out.push(fn(v)));
        return out;
    }
    filter(fn) {
        const out = new Stream();
        this.subscribe(v => { if (fn(v)) out.push(v); });
        return out;
    }
}
const s = new Stream();
const evens = s.filter(x => x % 2 === 0).map(x => x * 10);
const results = [];
evens.subscribe(v => results.push(v));
s.push(1, 2, 3, 4, 5);
console.log(results.join(","));
"#), vec!["20,40"]);
}

#[test]
fn debounce_throttle() {
    assert_eq!(run_js(r#"
function debounce(fn, delay) {
    let timer;
    return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), delay);
    };
}
function throttle(fn, limit) {
    let lastTime = 0;
    return (...args) => {
        const now = Date.now();
        if (now - lastTime >= limit) { lastTime = now; fn(...args); }
    };
}
// Verify they return functions
console.log(typeof debounce(() => {}, 100));
console.log(typeof throttle(() => {}, 100));
const calls = [];
const t = throttle(x => calls.push(x), 1000);
t(1); t(2); t(3);
console.log(calls.length);
"#), vec!["function", "function", "1"]);
}
