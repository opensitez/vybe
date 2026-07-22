use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class Private Fields (#field) Access & Encapsulation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_private_field_initialization_and_access() {
    let src = r#"
class Counter {
    #count = 0;
    increment() { this.#count++; }
    get value() { return this.#count; }
}
const c = new Counter();
c.increment();
c.increment();
console.log(c.value);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_class_private_field_outside_access_throws_syntaxerror() {
    let src = r#"
class Secret {
    #code = 1234;
}
const s = new Secret();
try {
    eval("s.#code");
} catch (e) {
    console.log("Outside Private Access Error");
}
"#;
    assert_eq!(run_js(src), vec!["Outside Private Access Error"]);
}

#[test]
fn test_js_class_private_field_initializer_expression_evaluation() {
    let src = r#"
let idCounter = 0;
class Item {
    #id = ++idCounter;
    getId() { return this.#id; }
}
const i1 = new Item();
const i2 = new Item();
console.log(`${i1.getId()}:${i2.getId()}`);
"#;
    assert_eq!(run_js(src), vec!["1:2"]);
}

#[test]
fn test_js_class_private_field_uninitialized_defaults_to_undefined() {
    let src = r#"
class Box {
    #content;
    getContent() { return this.#content === undefined; }
}
console.log(new Box().getContent());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_private_field_access_on_wrong_object_typeerror() {
    let src = r#"
class BankAccount {
    #balance = 100;
    getBalance(other) {
        return other.#balance; // Accessing #balance on non-BankAccount throws TypeError!
    }
}
const acc = new BankAccount();
try {
    acc.getBalance({});
} catch (e) {
    console.log("Private Access Wrong Receiver TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Private Access Wrong Receiver TypeError"]);
}

#[test]
fn test_js_class_private_field_cross_instance_access_same_class() {
    let src = r#"
class Vector {
    #x; #y;
    constructor(x, y) { this.#x = x; this.#y = y; }
    add(otherVector) {
        return new Vector(this.#x + otherVector.#x, this.#y + otherVector.#y);
    }
    toString() { return `(${this.#x},${this.#y})`; }
}
const v1 = new Vector(1, 2);
const v2 = new Vector(3, 4);
console.log(v1.add(v2).toString());
"#;
    assert_eq!(run_js(src), vec!["(4,6)"]);
}

#[test]
fn test_js_class_private_field_shadowing_subclass() {
    let src = r#"
class Base {
    #value = "BasePrivate";
    getBaseValue() { return this.#value; }
}
class Derived extends Base {
    #value = "DerivedPrivate"; // Distinct private field declaration!
    getDerivedValue() { return this.#value; }
}
const d = new Derived();
console.log(`${d.getBaseValue()}|${d.getDerivedValue()}`);
"#;
    assert_eq!(run_js(src), vec!["BasePrivate|DerivedPrivate"]);
}

#[test]
fn test_js_class_private_field_not_enumerable_in_keys() {
    let src = r#"
class User {
    #id = 1;
    name = "Alice";
}
const u = new User();
console.log(Object.keys(u).join(",") + "|Count=" + Object.getOwnPropertyNames(u).length);
"#;
    assert_eq!(run_js(src), vec!["name|Count=1"]);
}

#[test]
fn test_js_class_private_field_delete_operator_throws_syntaxerror() {
    let src = r#"
class Lock {
    #key = 99;
    tryDelete() {
        try {
            eval("delete this.#key;");
        } catch (e) {
            console.log("Delete Private Field Error");
        }
    }
}
new Lock().tryDelete();
"#;
    assert_eq!(run_js(src), vec!["Delete Private Field Error"]);
}

#[test]
fn test_js_class_private_field_this_in_initializer() {
    let src = r#"
class SelfRef {
    #self = this;
    getIsSelf() { return this.#self === this; }
}
console.log(new SelfRef().getIsSelf());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_class_private_field_arrow_function_closure_access() {
    let src = r#"
class Service {
    #token = "SECRET_TOKEN";
    getFetcher() {
        return () => `Bearer ${this.#token}`;
    }
}
const s = new Service();
const fetcher = s.getFetcher();
console.log(fetcher());
"#;
    assert_eq!(run_js(src), vec!["Bearer SECRET_TOKEN"]);
}

#[test]
fn test_js_class_private_field_computed_property_name_in_class_body() {
    let src = r#"
const propName = "dynamic";
class Component {
    #data = 500;
    [propName]() { return this.#data; }
}
console.log(new Component().dynamic());
"#;
    assert_eq!(run_js(src), vec!["500"]);
}

#[test]
fn test_js_class_private_field_super_constructor_initialization_order() {
    let src = r#"
const log = [];
class Base {
    constructor() {
        log.push("Base Ctor");
    }
}
class Derived extends Base {
    #field = (() => { log.push("Init Field"); return 10; })();
    constructor() {
        log.push("Before Super");
        super();
        log.push("After Super");
    }
}
new Derived();
console.log(log.join("->"));
"#;
    assert_eq!(
        run_js(src),
        vec!["Before Super->Base Ctor->Init Field->After Super"]
    );
}

#[test]
fn test_js_class_private_field_destructuring_assignment_inside_class() {
    let src = r#"
class DataHolder {
    #val = 100;
    swap(other) {
        [this.#val, other.#val] = [other.#val, this.#val];
    }
    getVal() { return this.#val; }
}
const d1 = new DataHolder();
const d2 = new DataHolder();
d1.swap(d2);
console.log(`${d1.getVal()}:${d2.getVal()}`);
"#;
    assert_eq!(run_js(src), vec!["100:100"]);
}

#[test]
fn test_js_class_private_field_nullish_coalescing_assignment() {
    let src = r#"
class Config {
    #cache = null;
    getCache() {
        this.#cache ??= "CachedData";
        return this.#cache;
    }
}
const cfg = new Config();
console.log(cfg.getCache());
"#;
    assert_eq!(run_js(src), vec!["CachedData"]);
}

#[test]
fn test_js_class_private_field_symbol_key_prohibited() {
    let src = r#"
try {
    eval("class Bad { #[Symbol('test')] = 1; }");
} catch (e) {
    console.log("Private Symbol Syntax Error");
}
"#;
    assert_eq!(run_js(src), vec!["Private Symbol Syntax Error"]);
}

#[test]
fn test_js_class_private_field_object_assign_does_not_copy_private_fields() {
    let src = r#"
class Container {
    #secret = 999;
    publicVal = 100;
    getSecret() { return this.#secret; }
}
const c1 = new Container();
const c2 = Object.assign({}, c1);
console.log(c2.publicVal + "|" + (typeof c2.getSecret === "undefined"));
"#;
    assert_eq!(run_js(src), vec!["100|true"]);
}

#[test]
fn test_js_class_private_field_getter_setter_interaction() {
    let src = r#"
class Range {
    #val = 0;
    set value(v) {
        if (v >= 0) this.#val = v;
    }
    get value() { return this.#val; }
}
const r = new Range();
r.value = 50;
r.value = -10; // Rejected by setter
console.log(r.value);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_class_private_field_multiple_declarations() {
    let src = r#"
class Node {
    #left; #right; #val;
    constructor(v) { this.#val = v; }
    setChildren(l, r) { this.#left = l; this.#right = r; }
    getSummary() {
        return `${this.#val}:[${this.#left.#val},${this.#right.#val}]`;
    }
}
const parent = new Node("Root");
parent.setChildren(new Node("L"), new Node("R"));
console.log(parent.getSummary());
"#;
    assert_eq!(run_js(src), vec!["Root:[L,R]"]);
}

#[test]
fn test_js_class_private_field_duplicate_declaration_throws_syntaxerror() {
    let src = r#"
try {
    eval("class Dup { #x = 1; #x = 2; }");
} catch (e) {
    console.log("Duplicate Private Field SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Duplicate Private Field SyntaxError"]);
}
