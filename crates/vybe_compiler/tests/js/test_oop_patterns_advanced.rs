/// Object-oriented programming patterns — advanced class features

use super::helpers::run_js;

#[test]
fn abstract_class_simulation() {
    assert_eq!(run_js(r#"
class Shape {
    area() { throw new Error("not implemented"); }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    area() { return Math.PI * this.r * this.r; }
}
const c = new Circle(1);
console.log(Math.abs(c.area() - Math.PI) < 0.0001);
let threw = false;
try { new Shape().area(); } catch { threw = true; }
console.log(threw);
"#), vec!["true", "true"]);
}

#[test]
fn interface_duck_typing() {
    assert_eq!(run_js(r#"
function implements_(obj, methods) {
    return methods.every(m => typeof obj[m] === "function");
}
const Serializable = ["serialize", "deserialize"];
class Config {
    serialize() { return JSON.stringify(this.data); }
    deserialize(s) { this.data = JSON.parse(s); return this; }
}
const cfg = new Config();
cfg.data = { x: 1 };
console.log(implements_(cfg, Serializable));
console.log(implements_({}, Serializable));
"#), vec!["true", "false"]);
}

#[test]
fn visitor_pattern() {
    assert_eq!(run_js(r#"
class NumberExpr { constructor(v) { this.v=v; } accept(visitor) { return visitor.visitNumber(this); } }
class AddExpr { constructor(l,r) { this.l=l; this.r=r; } accept(visitor) { return visitor.visitAdd(this); } }
class Evaluator {
    visitNumber(e) { return e.v; }
    visitAdd(e) { return e.l.accept(this) + e.r.accept(this); }
}
const expr = new AddExpr(new NumberExpr(3), new AddExpr(new NumberExpr(4), new NumberExpr(5)));
console.log(expr.accept(new Evaluator()));
"#), vec!["12"]);
}

#[test]
fn chain_of_responsibility() {
    assert_eq!(run_js(r#"
class Handler {
    constructor(next = null) { this.next = next; }
    handle(req) { return this.next ? this.next.handle(req) : "unhandled"; }
}
class AuthHandler extends Handler {
    handle(req) { return req.auth ? super.handle(req) : "unauthorized"; }
}
class RateLimitHandler extends Handler {
    handle(req) { return req.rate > 100 ? "rate limited" : super.handle(req); }
}
class ResourceHandler extends Handler {
    handle(req) { return "ok:" + req.resource; }
}
const chain = new AuthHandler(new RateLimitHandler(new ResourceHandler()));
console.log(chain.handle({ auth: true, rate: 50, resource: "data" }));
console.log(chain.handle({ auth: false, rate: 50, resource: "data" }));
console.log(chain.handle({ auth: true, rate: 200, resource: "data" }));
"#), vec!["ok:data", "unauthorized", "rate limited"]);
}

#[test]
fn template_method_pattern() {
    assert_eq!(run_js(r#"
class Report {
    generate() {
        return [this.header(), this.body(), this.footer()].join("|");
    }
    header() { return "Report"; }
    footer() { return "End"; }
    body() { throw new Error("override body"); }
}
class SalesReport extends Report {
    body() { return "Sales: 1000"; }
}
class HRReport extends Report {
    header() { return "HR Report"; }
    body() { return "Staff: 50"; }
}
console.log(new SalesReport().generate());
console.log(new HRReport().generate());
"#), vec!["Report|Sales: 1000|End", "HR Report|Staff: 50|End"]);
}

#[test]
fn lazy_initialization_class() {
    assert_eq!(run_js(r#"
class ExpensiveResource {
    #_data = null;
    get data() {
        if (!this.#_data) this.#_data = { computed: 42 };
        return this.#_data;
    }
}
const r = new ExpensiveResource();
console.log(r.data.computed);
console.log(r.data === r.data);
"#), vec!["42", "true"]);
}

#[test]
fn flyweight_pattern() {
    assert_eq!(run_js(r#"
class TreeType {
    constructor(name, color) { this.name=name; this.color=color; }
}
const _cache = new Map();
class TreeFactory {
    static get(name, color) {
        const key = name+color;
        if (!_cache.has(key)) _cache.set(key, new TreeType(name, color));
        return _cache.get(key);
    }
    static size() { return _cache.size; }
}
const t1 = TreeFactory.get("Oak", "green");
const t2 = TreeFactory.get("Oak", "green");
const t3 = TreeFactory.get("Pine", "dark");
console.log(t1 === t2);
console.log(t1 === t3);
console.log(TreeFactory.size());
"#), vec!["true", "false", "2"]);
}

#[test]
fn memento_pattern() {
    assert_eq!(run_js(r#"
class Editor {
    #content = "";
    #history = [];
    type(text) { this.#history.push(this.#content); this.#content += text; }
    undo() { if (this.#history.length) this.#content = this.#history.pop(); }
    get content() { return this.#content; }
}
const e = new Editor();
e.type("Hello");
e.type(" World");
console.log(e.content);
e.undo();
console.log(e.content);
e.undo();
console.log(e.content);
"#), vec!["Hello World", "Hello", ""]);
}

#[test]
fn composite_pattern() {
    assert_eq!(run_js(r#"
class File {
    constructor(name, size) { this.name=name; this.size=size; }
    totalSize() { return this.size; }
}
class Folder {
    constructor(name) { this.name=name; this.children=[]; }
    add(child) { this.children.push(child); return this; }
    totalSize() { return this.children.reduce((s,c)=>s+c.totalSize(),0); }
}
const root = new Folder("root")
    .add(new File("a.txt", 100))
    .add(new Folder("sub").add(new File("b.txt", 200)).add(new File("c.txt", 300)));
console.log(root.totalSize());
"#), vec!["600"]);
}

#[test]
fn proxy_validation_pattern() {
    assert_eq!(run_js(r#"
function createValidated(target, validators) {
    return new Proxy(target, {
        set(obj, prop, value) {
            if (validators[prop] && !validators[prop](value)) throw new Error(`Invalid ${prop}`);
            obj[prop] = value;
            return true;
        }
    });
}
const person = createValidated({}, { age: v => typeof v === "number" && v >= 0 && v <= 150 });
person.name = "Alice";
person.age = 30;
console.log(person.name);
console.log(person.age);
let threw = false;
try { person.age = -5; } catch { threw = true; }
console.log(threw);
"#), vec!["Alice", "30", "true"]);
}

#[test]
fn event_sourcing_pattern() {
    assert_eq!(run_js(r#"
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
"#), vec!["Alice", "30"]);
}
