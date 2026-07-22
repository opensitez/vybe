use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Class Getter/Setter Overriding & Super Access
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_class_override_getter_only() {
    let src = r#"
class Base {
    get val() { return 10; }
}
class Derived extends Base {
    get val() { return super.val * 5; }
}
console.log(new Derived().val);
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_class_override_setter_only() {
    let src = r#"
class Base {
    set val(v) { this._val = v + 1; }
    get val() { return this._val; }
}
class Derived extends Base {
    set val(v) { super.val = v * 10; }
}
const d = new Derived();
d.val = 5;
console.log(d.val);
"#;
    assert_eq!(run_js(src), vec!["51"]);
}

#[test]
fn test_js_class_override_getter_and_setter_both() {
    let src = r#"
class Base {
    _score = 0;
    get score() { return this._score; }
    set score(v) { this._score = v; }
}
class Derived extends Base {
    get score() { return super.score + 100; }
    set score(v) { super.score = v * 2; }
}
const d = new Derived();
d.score = 10;
console.log(d.score); // (10 * 2) + 100 = 120
"#;
    assert_eq!(run_js(src), vec!["120"]);
}

#[test]
fn test_js_class_override_getter_preserves_backing_field_per_instance() {
    let src = r#"
class Base {
    constructor(v) { this._v = v; }
    get value() { return this._v; }
}
class Derived extends Base {
    get value() { return `[${super.value}]`; }
}
const d1 = new Derived("A");
const d2 = new Derived("B");
console.log(`${d1.value}:${d2.value}`);
"#;
    assert_eq!(run_js(src), vec!["[A]:[B]"]);
}

#[test]
fn test_js_class_static_getter_setter_overriding() {
    let src = r#"
class Base {
    static get env() { return "base"; }
}
class Derived extends Base {
    static get env() { return super.env.toUpperCase(); }
}
console.log(Derived.env);
"#;
    assert_eq!(run_js(src), vec!["BASE"]);
}

#[test]
fn test_js_class_getter_override_without_super_hides_parent_getter() {
    let src = r#"
class Parent {
    get name() { return "ParentName"; }
}
class Child extends Parent {
    get name() { return "ChildName"; }
}
console.log(new Child().name);
"#;
    assert_eq!(run_js(src), vec!["ChildName"]);
}

#[test]
fn test_js_class_setter_override_without_super_hides_parent_setter() {
    let src = r#"
class Parent {
    set data(v) { this.parentData = v; }
}
class Child extends Parent {
    set data(v) { this.childData = v; }
}
const c = new Child();
c.data = "Test";
console.log(c.parentData + "|" + c.childData);
"#;
    assert_eq!(run_js(src), vec!["undefined|Test"]);
}

#[test]
fn test_js_class_override_getter_in_subclass_replaces_parent_value_property() {
    let src = r#"
class Parent {
    title = "ParentTitle";
}
class Child extends Parent {
    get title() { return "ChildGetterTitle"; }
}
console.log(new Child().title);
"#;
    assert_eq!(run_js(src), vec!["ParentTitle"]); // Instance field 'title' set in constructor shadows prototype getter!
}

#[test]
fn test_js_class_override_value_property_in_subclass_shadows_parent_getter() {
    let src = r#"
class Parent {
    get title() { return "ParentTitleGetter"; }
}
class Child extends Parent {
    constructor() {
        super();
        this.title = "ChildInstanceField";
    }
}
console.log(new Child().title);
"#;
    assert_eq!(run_js(src), vec!["ChildInstanceField"]);
}

#[test]
fn test_js_class_override_getter_returning_object() {
    let src = r#"
class Base {
    get config() { return { port: 80 }; }
}
class Derived extends Base {
    get config() {
        const baseCfg = super.config;
        return { ...baseCfg, ssl: true };
    }
}
const d = new Derived();
console.log(`${d.config.port}:${d.config.ssl}`);
"#;
    assert_eq!(run_js(src), vec!["80:true"]);
}

#[test]
fn test_js_class_override_getter_validation_throws() {
    let src = r#"
class Base {
    get age() { return this._age; }
    set age(v) { this._age = v; }
}
class Derived extends Base {
    set age(v) {
        if (v < 0) throw new RangeError("Negative Age Error");
        super.age = v;
    }
}
const d = new Derived();
d.age = 20;
console.log(d.age);
try {
    d.age = -5;
} catch (e) {
    console.log("RangeError Caught");
}
"#;
    assert_eq!(run_js(src), vec!["20", "RangeError Caught"]);
}

#[test]
fn test_js_class_override_computed_getter_name() {
    let src = r#"
const propKey = "dynamicKey";
class Base {
    get [propKey]() { return "BaseVal"; }
}
class Derived extends Base {
    get [propKey]() { return super[propKey] + "_Derived"; }
}
console.log(new Derived().dynamicKey);
"#;
    assert_eq!(run_js(src), vec!["BaseVal_Derived"]);
}

#[test]
fn test_js_class_override_symbol_getter() {
    let src = r#"
const sym = Symbol("getterKey");
class Base {
    get [sym]() { return 100; }
}
class Derived extends Base {
    get [sym]() { return super[sym] * 3; }
}
console.log(new Derived()[sym]);
"#;
    assert_eq!(run_js(src), vec!["300"]);
}

#[test]
fn test_js_class_override_getter_in_object_assign_copy() {
    let src = r#"
class Base {
    get id() { return 123; }
}
class Derived extends Base {
    get id() { return super.id + 1; }
}
const d = new Derived();
const copy = Object.assign({}, d);
console.log(copy.id);
"#;
    assert_eq!(run_js(src), vec!["124"]);
}

#[test]
fn test_js_class_override_getter_side_effect_count() {
    let src = r#"
let reads = 0;
class Base {
    get count() { reads++; return 1; }
}
class Derived extends Base {
    get count() { return super.count + super.count; }
}
console.log(new Derived().count + "|Reads=" + reads);
"#;
    assert_eq!(run_js(src), vec!["2|Reads=2"]);
}

#[test]
fn test_js_class_override_setter_returns_value_in_assignment_chain() {
    let src = r#"
class Base {
    set x(v) { this._x = v; }
    get x() { return this._x; }
}
class Derived extends Base {
    set x(v) { super.x = v; }
}
const d = new Derived();
const assigned = (d.x = 99);
console.log(assigned + "|" + d.x);
"#;
    assert_eq!(run_js(src), vec!["99|99"]);
}

#[test]
fn test_js_class_override_getter_with_private_field_backing() {
    let src = r#"
class Base {
    #val = 50;
    get val() { return this.#val; }
    set val(v) { this.#val = v; }
}
class Derived extends Base {
    get val() { return super.val * 2; }
}
const d = new Derived();
d.val = 100;
console.log(d.val);
"#;
    assert_eq!(run_js(src), vec!["200"]);
}

#[test]
fn test_js_class_override_getter_descriptor_inspection() {
    let src = r#"
class Base {
    get item() { return "B"; }
}
class Derived extends Base {
    get item() { return "D"; }
}
const desc = Object.getOwnPropertyDescriptor(Derived.prototype, "item");
console.log(typeof desc.get + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["function|false|true"]);
}

#[test]
fn test_js_class_override_getter_async_promise() {
    let src = r#"
class Base {
    get asyncVal() { return Promise.resolve(10); }
}
class Derived extends Base {
    get asyncVal() {
        return super.asyncVal.then(v => v * 4);
    }
}
new Derived().asyncVal.then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["40"]);
}

#[test]
fn test_js_class_override_getter_generator() {
    let src = r#"
class Base {
    get stream() {
        return function*() { yield 1; yield 2; };
    }
}
class Derived extends Base {
    get stream() {
        const baseStream = super.stream;
        return function*() {
            yield* baseStream();
            yield 3;
        };
    }
}
console.log([...new Derived().stream()].join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,3"]);
}
