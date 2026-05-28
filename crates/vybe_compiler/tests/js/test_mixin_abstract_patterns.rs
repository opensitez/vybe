/// Mixin patterns, abstract classes, multiple inheritance via mixins

use super::helpers::run_js;

#[test]
fn simple_mixin_adds_methods() {
    assert_eq!(run_js(r#"
const Flyable = (Base) => class extends Base {
    fly() { return this.name + " is flying"; }
};
class Animal {
    constructor(name) { this.name = name; }
}
class Bird extends Flyable(Animal) {}
const b = new Bird("Eagle");
console.log(b.fly());
console.log(b instanceof Animal);
"#), vec!["Eagle is flying", "true"]);
}

#[test]
fn mixin_chain_two_mixins() {
    assert_eq!(run_js(r#"
const Swimmable = Base => class extends Base {
    swim() { return this.name + " swims"; }
};
const Flyable = Base => class extends Base {
    fly() { return this.name + " flies"; }
};
class Animal {
    constructor(name) { this.name = name; }
}
class Duck extends Swimmable(Flyable(Animal)) {}
const d = new Duck("Donald");
console.log(d.swim());
console.log(d.fly());
"#), vec!["Donald swims", "Donald flies"]);
}

#[test]
fn abstract_class_pattern_with_new_target() {
    assert_eq!(run_js(r#"
class Abstract {
    constructor() {
        if (new.target === Abstract) throw new Error("Cannot instantiate Abstract");
    }
    method() { throw new Error("Not implemented"); }
}
class Concrete extends Abstract {
    method() { return "implemented"; }
}
let threw = false;
try { new Abstract(); } catch { threw = true; }
console.log(threw);
console.log(new Concrete().method());
"#), vec!["true", "implemented"]);
}

#[test]
fn mixin_with_super() {
    assert_eq!(run_js(r#"
const Timestamped = Base => class extends Base {
    constructor(...args) {
        super(...args);
        this.createdAt = 0;
    }
};
class User {
    constructor(name) { this.name = name; }
}
class TimestampedUser extends Timestamped(User) {}
const u = new TimestampedUser("Alice");
console.log(u.name);
console.log(u.createdAt);
"#), vec!["Alice", "0"]);
}

#[test]
fn mixin_preserves_instanceof_chain() {
    assert_eq!(run_js(r#"
const Mixin = Base => class extends Base {};
class Root {}
class Child extends Mixin(Root) {}
const c = new Child();
console.log(c instanceof Child);
console.log(c instanceof Root);
"#), vec!["true", "true"]);
}

#[test]
fn functional_mixin_copies_methods() {
    assert_eq!(run_js(r#"
function applyMixin(target, mixin) {
    Object.assign(target.prototype, mixin);
}
const Serializable = {
    serialize() { return JSON.stringify(this); }
};
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
applyMixin(Point, Serializable);
const p = new Point(1, 2);
console.log(p.serialize());
"#), vec!["{\"x\":1,\"y\":2}"]);
}

#[test]
fn class_with_private_mixin_field() {
    assert_eq!(run_js(r#"
class Counter {
    #count = 0;
    increment() { this.#count++; }
    getCount() { return this.#count; }
}
const c = new Counter();
c.increment();
c.increment();
console.log(c.getCount());
"#), vec!["2"]);
}

#[test]
fn super_method_call_in_mixin() {
    assert_eq!(run_js(r#"
const Logger = Base => class extends Base {
    greet() {
        return "[LOG] " + super.greet();
    }
};
class Person {
    constructor(name) { this.name = name; }
    greet() { return "Hi, I'm " + this.name; }
}
class LoggedPerson extends Logger(Person) {}
const p = new LoggedPerson("Bob");
console.log(p.greet());
"#), vec!["[LOG] Hi, I'm Bob"]);
}
