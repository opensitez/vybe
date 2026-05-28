/// Class static features — static fields, methods, init blocks, private static

use super::helpers::run_js;

#[test]
fn static_field_shared_across_instances() {
    assert_eq!(run_js(r#"
class Counter {
    static count = 0;
    constructor() { Counter.count++; }
}
new Counter();
new Counter();
new Counter();
console.log(Counter.count);
"#), vec!["3"]);
}

#[test]
fn static_method_called_on_class() {
    assert_eq!(run_js(r#"
class MathHelper {
    static add(a, b) { return a + b; }
    static multiply(a, b) { return a * b; }
}
console.log(MathHelper.add(3, 4));
console.log(MathHelper.multiply(3, 4));
"#), vec!["7", "12"]);
}

#[test]
fn static_method_not_on_instance() {
    assert_eq!(run_js(r#"
class Foo {
    static bar() { return 42; }
}
const f = new Foo();
console.log(typeof f.bar);
console.log(Foo.bar());
"#), vec!["undefined", "42"]);
}

#[test]
fn static_initializer_block() {
    assert_eq!(run_js(r#"
class Config {
    static values;
    static {
        Config.values = [1, 2, 3];
        Config.sum = Config.values.reduce((a, b) => a + b, 0);
    }
}
console.log(Config.sum);
console.log(Config.values.join(","));
"#), vec!["6", "1,2,3"]);
}

#[test]
fn static_private_field() {
    assert_eq!(run_js(r#"
class Registry {
    static #instances = 0;
    static create() {
        Registry.#instances++;
        return new Registry();
    }
    static getCount() { return Registry.#instances; }
}
Registry.create();
Registry.create();
console.log(Registry.getCount());
"#), vec!["2"]);
}

#[test]
fn static_private_not_accessible_outside() {
    assert_eq!(run_js(r##"
class Foo {
    static #secret = 42;
    static get() { return Foo.#secret; }
}
console.log(Foo.get());
const key = "#" + "secret";
console.log(Foo[key] === undefined);
"##), vec!["42", "true"]);
}

#[test]
fn static_field_initializer_order() {
    assert_eq!(run_js(r#"
const log = [];
class Foo {
    static a = (log.push("a"), 1);
    static b = (log.push("b"), 2);
    static c = (log.push("c"), 3);
}
console.log(log.join(","));
"#), vec!["a,b,c"]);
}

#[test]
fn static_method_inherited_by_subclass() {
    assert_eq!(run_js(r#"
class Animal {
    static describe() { return "I am " + this.name; }
}
class Dog extends Animal {}
console.log(Dog.describe());
"#), vec!["I am Dog"]);
}

#[test]
fn static_this_in_initializer() {
    assert_eq!(run_js(r#"
class Foo {
    static x = 10;
    static y = this.x * 2;
}
console.log(Foo.y);
"#), vec!["20"]);
}

#[test]
fn static_init_block_can_access_private() {
    assert_eq!(run_js(r#"
class Secret {
    static #value;
    static {
        Secret.#value = 42;
    }
    static get() { return Secret.#value; }
}
console.log(Secret.get());
"#), vec!["42"]);
}
