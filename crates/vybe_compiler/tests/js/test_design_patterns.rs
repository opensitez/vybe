/// Design patterns in JavaScript — Observer, Builder, Strategy, Command, State
use super::helpers::run_js;

#[test]
fn observer_pattern() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["1,2"]
    );
}

#[test]
fn builder_pattern() {
    assert_eq!(
        run_js(
            r#"
class QueryBuilder {
    constructor() { this._table = ""; this._where = []; this._limit = null; }
    from(table) { this._table = table; return this; }
    where(cond) { this._where.push(cond); return this; }
    limit(n) { this._limit = n; return this; }
    build() {
        let q = `SELECT * FROM ${this._table}`;
        if (this._where.length) q += ` WHERE ${this._where.join(" AND ")}`;
        if (this._limit) q += ` LIMIT ${this._limit}`;
        return q;
    }
}
const query = new QueryBuilder()
    .from("users")
    .where("age > 18")
    .where("active = 1")
    .limit(10)
    .build();
console.log(query);
"#
        ),
        vec!["SELECT * FROM users WHERE age > 18 AND active = 1 LIMIT 10"]
    );
}

#[test]
fn strategy_pattern() {
    assert_eq!(
        run_js(
            r#"
class Sorter {
    constructor(strategy) { this.strategy = strategy; }
    sort(arr) { return this.strategy([...arr]); }
}
const ascending = arr => arr.sort((a, b) => a - b);
const descending = arr => arr.sort((a, b) => b - a);
const nums = [3, 1, 4, 1, 5, 9, 2, 6];
const asc = new Sorter(ascending);
const desc = new Sorter(descending);
console.log(asc.sort(nums).join(","));
console.log(desc.sort(nums).join(","));
"#
        ),
        vec!["1,1,2,3,4,5,6,9", "9,6,5,4,3,2,1,1"]
    );
}

#[test]
fn command_pattern() {
    assert_eq!(
        run_js(
            r#"
class TextEditor {
    constructor() { this.text = ""; this.history = []; }
    execute(command) {
        this.text = command.execute(this.text);
        this.history.push(command);
    }
    undo() {
        const command = this.history.pop();
        if (command) this.text = command.undo(this.text);
    }
}
const append = text => ({
    execute: current => current + text,
    undo: current => current.slice(0, -text.length)
});
const editor = new TextEditor();
editor.execute(append("Hello"));
editor.execute(append(" World"));
console.log(editor.text);
editor.undo();
console.log(editor.text);
"#
        ),
        vec!["Hello World", "Hello"]
    );
}

#[test]
fn singleton_pattern() {
    assert_eq!(
        run_js(
            r#"
class Config {
    static #instance = null;
    #settings = {};
    static getInstance() {
        if (!Config.#instance) Config.#instance = new Config();
        return Config.#instance;
    }
    set(key, val) { this.#settings[key] = val; return this; }
    get(key) { return this.#settings[key]; }
}
const a = Config.getInstance();
const b = Config.getInstance();
a.set("theme", "dark");
console.log(a === b);
console.log(b.get("theme"));
"#
        ),
        vec!["true", "dark"]
    );
}

#[test]
fn decorator_pattern_functional() {
    assert_eq!(
        run_js(
            r#"
function withLogging(fn, name) {
    return function(...args) {
        const result = fn(...args);
        console.log(`${name}(${args}) = ${result}`);
        return result;
    };
}
const add = withLogging((a, b) => a + b, "add");
add(3, 4);
"#
        ),
        vec!["add(3,4) = 7"]
    );
}

#[test]
fn state_machine_pattern() {
    assert_eq!(
        run_js(
            r#"
class TrafficLight {
    #state = "red";
    next() {
        const transitions = { red: "green", green: "yellow", yellow: "red" };
        this.#state = transitions[this.#state];
        return this.#state;
    }
    get state() { return this.#state; }
}
const light = new TrafficLight();
console.log(light.state);
console.log(light.next());
console.log(light.next());
console.log(light.next());
"#
        ),
        vec!["red", "green", "yellow", "red"]
    );
}

#[test]
fn iterator_pattern_custom() {
    assert_eq!(
        run_js(
            r#"
class Range {
    constructor(start, end, step = 1) {
        this.start = start; this.end = end; this.step = step;
    }
    [Symbol.iterator]() {
        let cur = this.start;
        const { end, step } = this;
        return {
            next() {
                return cur < end
                    ? { value: cur, done: false, ...{ next: () => { cur += step; } } }
                    : { done: true };
            }
        };
    }
}
// Use a generator instead for cleaner implementation
function* range(start, end, step = 1) {
    for (let i = start; i < end; i += step) yield i;
}
console.log([...range(0, 10, 2)].join(","));
console.log([...range(1, 6)].join(","));
"#
        ),
        vec!["0,2,4,6,8", "1,2,3,4,5"]
    );
}
