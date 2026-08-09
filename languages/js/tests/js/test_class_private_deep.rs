/// Class private class fields — #field, #method, brand checking, accessor
use super::helpers::run_js;

#[test]
fn private_field_inaccessible_outside() {
    assert_eq!(
        run_js(
            r##"
class Wallet {
    #balance = 0;
    deposit(n) { this.#balance += n; }
    get balance() { return this.#balance; }
}
const w = new Wallet();
w.deposit(100);
console.log(w.balance);
const key = "#" + "balance";
console.log(w[key] === undefined);
"##
        ),
        vec!["100", "true"]
    );
}

#[test]
fn private_method() {
    assert_eq!(
        run_js(
            r#"
class Validator {
    #validate(x) { return x > 0; }
    check(x) { return this.#validate(x) ? "valid" : "invalid"; }
}
const v = new Validator();
console.log(v.check(5));
console.log(v.check(-1));
"#
        ),
        vec!["valid", "invalid"]
    );
}

#[test]
fn private_static_field() {
    assert_eq!(
        run_js(
            r#"
class IdGenerator {
    static #nextId = 1;
    static generate() { return IdGenerator.#nextId++; }
}
console.log(IdGenerator.generate());
console.log(IdGenerator.generate());
console.log(IdGenerator.generate());
"#
        ),
        vec!["1", "2", "3"]
    );
}

#[test]
fn private_field_per_instance() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    #count = 0;
    inc() { this.#count++; }
    get() { return this.#count; }
}
const a = new Counter();
const b = new Counter();
a.inc(); a.inc();
b.inc();
console.log(a.get());
console.log(b.get());
"#
        ),
        vec!["2", "1"]
    );
}

#[test]
fn private_field_in_operator_brand_check() {
    assert_eq!(
        run_js(
            r#"
class Foo {
    #x;
    static isFoo(obj) { return #x in obj; }
}
const f = new Foo();
console.log(Foo.isFoo(f));
console.log(Foo.isFoo({}));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn private_accessor_auto_accessor() {
    assert_eq!(
        run_js(
            r#"
class Temperature {
    #celsius = 0;
    get celsius() { return this.#celsius; }
    set celsius(v) { this.#celsius = v; }
    get fahrenheit() { return this.#celsius * 9/5 + 32; }
}
const t = new Temperature();
t.celsius = 100;
console.log(t.celsius);
console.log(t.fahrenheit);
"#
        ),
        vec!["100", "212"]
    );
}

#[test]
fn private_fields_not_inherited() {
    assert_eq!(
        run_js(
            r#"
class Parent {
    #x = 42;
    getX() { return this.#x; }
}
class Child extends Parent {
    // Cannot access Parent's #x directly
    getFromParent() { return this.getX(); }
}
const c = new Child();
console.log(c.getFromParent());
"#
        ),
        vec!["42"]
    );
}

#[test]
fn private_method_calling_private_method() {
    assert_eq!(
        run_js(
            r#"
class Parser {
    #tokenize(str) { return str.split(" "); }
    #process(tokens) { return tokens.map(t => t.toUpperCase()); }
    parse(str) { return this.#process(this.#tokenize(str)).join(","); }
}
const p = new Parser();
console.log(p.parse("hello world foo"));
"#
        ),
        vec!["HELLO,WORLD,FOO"]
    );
}

#[test]
fn static_private_with_instance_methods() {
    assert_eq!(
        run_js(
            r#"
class EventLogger {
    static #log = [];
    static getLog() { return [...EventLogger.#log]; }
    log(event) { EventLogger.#log.push(event); }
}
const logger = new EventLogger();
logger.log("start");
logger.log("process");
logger.log("end");
console.log(EventLogger.getLog().join(","));
"#
        ),
        vec!["start,process,end"]
    );
}

#[test]
fn test_private_static_method_in_static_getter() {
    assert_eq!(
        run_js(
            r#"
class Secret {
    static #compute() { return "staticSecret"; }
    static get secret() { return Secret.#compute(); }
}
console.log(Secret.secret);
"#
        ),
        vec!["staticSecret"]
    );
}
