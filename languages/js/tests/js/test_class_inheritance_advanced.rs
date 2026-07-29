/// Class inheritance edge cases — super in static, extends expression, new.target
use super::helpers::run_js;

#[test]
fn super_in_constructor_chain() {
    assert_eq!(
        run_js(
            r#"
class A {
    constructor(x) { this.x = x; }
}

class B extends A {
    constructor(x, y) {
        super(x);
        this.y = y;
    }
}
class C extends B {
    constructor(x, y, z) {
        super(x, y);
        this.z = z;
    }
}
const c = new C(1, 2, 3);
console.log(c.x);
console.log(c.y);
console.log(c.z);
"#
        ),
        vec!["1", "2", "3"]
    );
}

#[test]
fn super_getter_uses_base_implementation() {
    let src = r#"
class Base {
    get label() {
        return "base";
    }
}

class Child extends Base {
    constructor(tag) {
        super();
        this.tag = tag;
    }

    get label() {
        return super.label + "|" + this.tag;
    }
}

const c = new Child("node");
console.log(c.label);
console.log(c instanceof Base);
"#;
    assert_eq!(run_js(src), vec!["base|node", "true"]);
}

#[test]
fn instance_field_initializer_uses_super_method() {
    let src = r#"
class Base {
    baseLabel() {
        return "base";
    }
}

class Child extends Base {
    label = super.baseLabel();
}

console.log(new Child().label);
"#;
    assert_eq!(run_js(src), vec!["base"]);
}

#[test]
fn super_method_in_static_context() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    static describe() { return "Animal"; }
}
class Dog extends Animal {
    static describe() { return super.describe() + "/Dog"; }
}
console.log(Dog.describe());
"#
        ),
        vec!["Animal/Dog"]
    );
}

#[test]
fn new_target_is_subclass_in_super_constructor() {
    assert_eq!(
        run_js(
            r#"
class Base {
    constructor() {
        this.constructedAs = this.constructor.name;
    }
}
class Derived extends Base {}
const b = new Base();
const d = new Derived();
console.log(b.constructedAs);
console.log(d.constructedAs);
"#
        ),
        vec!["Base", "Derived"]
    );
}

#[test]
fn new_target_reflects_constructed_class_in_inheritance_chain() {
    assert_eq!(
        run_js(
            r#"
class Base {
    constructor() {
        this.requested = new.target;
    }
}
class Child extends Base {}
class GrandChild extends Child {}
console.log(new Child().requested.name);
console.log(new GrandChild().requested.name);
"#
        ),
        vec!["Child", "GrandChild"]
    );
}

#[test]
fn constructor_access_to_this_before_super_throws() {
    assert_eq!(
        run_js(
            r#"
class Base {
    constructor() {
        this.base = true;
    }
}
class Child extends Base {
    constructor() {
        try {
            this.beforeSuper = true;
        } catch (e) {
            console.log(e.name);
            return;
        }
        super();
    }
}
new Child();
"#
        ),
        vec!["ReferenceError"]
    );
}

#[test]
fn extends_with_expression() {
    assert_eq!(
        run_js(
            r#"
function makeBase(msg) {
    return class {
        greet() { return msg; }
    };
}
const Base = makeBase("hello from factory");
class Derived extends Base {}
const d = new Derived();
console.log(d.greet());
"#
        ),
        vec!["hello from factory"]
    );
}

#[test]
fn subclass_overrides_getter() {
    assert_eq!(
        run_js(
            r#"
class Shape {
    get area() { return 0; }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    get area() { return Math.PI * this.r * this.r; }
}
const c = new Circle(1);
console.log(c.area.toFixed(5));
"#
        ),
        vec!["3.14159"]
    );
}

#[test]
fn parent_method_accessible_via_super() {
    assert_eq!(
        run_js(
            r#"
class Logger {
    log(msg) { return "[LOG] " + msg; }
}
class PrefixLogger extends Logger {
    constructor(prefix) {
        super();
        this.prefix = prefix;
    }
    log(msg) {
        return super.log(this.prefix + ": " + msg);
    }
}
const logger = new PrefixLogger("App");
console.log(logger.log("started"));
"#
        ),
        vec!["[LOG] App: started"]
    );
}

#[test]
fn class_in_expression_position() {
    assert_eq!(
        run_js(
            r#"
const Greeter = class NamedGreeter {
    greet(name) { return "Hello " + name; }
};
const g = new Greeter();
console.log(g.greet("World"));
console.log(typeof g.greet);
"#
        ),
        vec!["Hello World", "function"]
    );
}

#[test]
fn computed_super_method_call_in_instance_method() {
    let src = r#"
class Base {
    label() {
        return "base-label";
    }
}

class Child extends Base {
    getLabel() {
        const key = "label";
        return super[key]();
    }
}

const c = new Child();
console.log(c.getLabel());
"#;
    assert_eq!(run_js(src), vec!["base-label"]);
}

#[test]
fn computed_super_static_getter_call() {
    let src = r#"
class Base {
    static get marker() {
        return "base-marker";
    }
}

class Child extends Base {
    static getComputedMarker() {
        const key = "marker";
        return super[key];
    }
}

console.log(Child.getComputedMarker());
"#;
    assert_eq!(run_js(src), vec!["base-marker"]);
}

#[test]
fn extends_null_sets_prototype_to_null() {
    let src = r#"
class NullPrototype extends null {}
console.log(Object.getPrototypeOf(NullPrototype.prototype) === null);
console.log(Object.getPrototypeOf(NullPrototype) === Function.prototype);
"#;
    assert_eq!(run_js(src), vec!["false", "true"]);
}

#[test]
fn instanceof_in_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
class A {}
class B extends A {}
class C extends B {}
const c = new C();
console.log(c instanceof C);
console.log(c instanceof B);
console.log(c instanceof A);
console.log(c instanceof Object);
const b = new B();
console.log(b instanceof C); // false — b is not a C
"#
        ),
        vec!["true", "true", "true", "true", "false"]
    );
}

#[test]
fn subclass_calls_super_method_with_this() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    constructor() { this.count = 0; }
    increment() { this.count++; return this; }
}
class BoundedCounter extends Counter {
    constructor(max) {
        super();
        this.max = max;
    }
    increment() {
        if (this.count < this.max) super.increment();
        return this;
    }
}
const bc = new BoundedCounter(3);
bc.increment().increment().increment().increment().increment();
console.log(bc.count);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn inherits_prototype_methods() {
    assert_eq!(
        run_js(
            r#"
class EventEmitter {
    constructor() { this._handlers = {}; }
    on(event, fn) {
        (this._handlers[event] = this._handlers[event] || []).push(fn);
    }
    emit(event, ...args) {
        (this._handlers[event] || []).forEach(fn => fn(...args));
    }
}
class Button extends EventEmitter {
    click() { this.emit("click", this); }
}
const btn = new Button();
const log = [];
btn.on("click", () => log.push("clicked"));
btn.click();
btn.click();
console.log(log.join(","));
"#
        ),
        vec!["clicked,clicked"]
    );
}

#[test]
fn instance_fields_are_initialized_after_super() {
    let src = r#"
class Base {
    baseField = "base";
    constructor() { this.baseFlag = true; }
}
class Child extends Base {
    childField = "child";
    constructor() {
        super();
        this.childFlag = true;
    }
}
const c = new Child();
console.log(c.baseField + "|" + c.childField + "|" + c.baseFlag + "|" + c.childFlag);
"#;
    assert_eq!(run_js(src), vec!["base|child|true|true"]);
}

#[test]
fn prototype_chain_is_connected_for_base_and_child_prototypes() {
    let src = r#"
class Base {}
class Child extends Base {}
console.log(Object.getPrototypeOf(Child.prototype) === Base.prototype);
console.log(Object.getPrototypeOf(Child) === Base);
const c = new Child();
console.log(c instanceof Base);
"#;
    assert_eq!(run_js(src), vec!["true", "true", "true"]);
}

#[test]
fn static_members_inherited_and_overridden() {
    let src = r#"
class Base {
    static version() { return "base"; }
    static metadata() { return "Base:" + this.name + ":" + this.version(); }
}

class Child extends Base {
    static version() { return "child"; }
    static metadataFromSuper() { return super.version() + "|" + super.metadata(); }
}

console.log(Object.getPrototypeOf(Child) === Base);
console.log(Child.version());
console.log(Child.metadataFromSuper());
"#;
    assert_eq!(run_js(src), vec!["true", "child", "base|Base:Child:child"]);
}

#[test]
fn inherited_super_accessor_chain() {
    assert_eq!(
        run_js(
            r#"
class Base {
    set value(raw) {
        this._value = `base:${raw}`;
    }
    get value() {
        return this._value;
    }
}

class Child extends Base {
    set value(raw) {
        super.value = `child:${raw}`;
    }
    get value() {
        return `child->${super.value}`;
    }
}

const c = new Child();
c.value = "payload";
console.log(c.value);
console.log(c._value);
"#
        ),
        vec!["child->base:child:payload", "base:child:payload"]
    );
}

#[test]
fn static_super_member_accesses_parent_field() {
    assert_eq!(
        run_js(
            r#"
class Base {
    static level = "base";
}

class Child extends Base {
    static level = "child";
    static describe() {
        return `${super.level}:${this.level}`;
    }
}
console.log(Child.describe());
"#
        ),
        vec!["base:child"]
    );
}

#[test]
fn static_and_instance_properties_show_expected_resolution_order() {
    assert_eq!(
        run_js(
            r#"
class Base {
    static namespace = "base";
    static get label() { return `static:${this.namespace}`; }
    marker = "base";
    label() { return "base"; }
}

class Child extends Base {
    static namespace = "child";
    static get fullLabel() {
        return `${super.namespace}|${this.namespace}|${super.label}`;
    }
    marker = "child";
    label() { return `child:${super.label()}`; }
}

const c = new Child();
console.log(c.marker);
console.log(c.label());
console.log(Child.namespace);
console.log(Child.fullLabel);
console.log(Child.label); // method name, not executed
"#
        ),
        vec!["child", "child:base", "child", "base|child|static:child", "static:child"]
    );
}

#[test]
fn constructor_returning_object_override_skips_instance_initialization() {
    let src = r#"
class Base {
    constructor() {
        this.tag = "base";
    }
}

class Child extends Base {
    constructor() {
        super();
        return { tag: "replacement", marker: "overridden" };
    }
}

const c = new Child();
console.log(c.tag);
console.log(c.marker);
console.log(c instanceof Base);
"#;
    assert_eq!(run_js(src), vec!["replacement", "overridden", "false"]);
}

#[test]
fn constructor_returning_null_keeps_instance() {
    let src = r#"
class Base {
    constructor() {
        this.tag = "base";
    }
}

class Child extends Base {
    constructor() {
        super();
        this.extra = "child";
        return null;
    }
}

const c = new Child();
console.log(c instanceof Base);
console.log(c.tag);
console.log(c.extra);
"#;
    assert_eq!(run_js(src), vec!["true", "base", "child"]);
}

#[test]
fn base_prototype_for_instance_fields_after_subclass_fields() {
    let src = r#"
class Base {
    baseField = "base";
    constructor() {}
}
class Child extends Base {
    childField = "child";
    constructor() {
        super();
        this.baseAndChild = this.baseField + "|" + this.childField;
    }
}
const c = new Child();
console.log(c.baseAndChild);
"#;
    assert_eq!(run_js(src), vec!["base|child"]);
}

#[test]
fn instance_field_initializer_uses_base_getter_via_super() {
    let src = r#"
class Base {
    get name() {
        return "base";
    }
}

class Child extends Base {
    label = super.name;
    readLabel() { return super.name; }
}

const c = new Child();
console.log(c.label);
console.log(c.readLabel());
"#;
    assert_eq!(
        run_js(src),
        vec!["base", "base"]
    );
}

#[test]
fn test_super_property_assignment_targets_receiver_this() {
    let src = r#"
class Base {}
class Child extends Base {
    setProp(v) {
        super.x = v;
    }
}
const c = new Child();
c.setProp(42);
console.log(c.x + "|" + ("x" in Base.prototype));
"#;
    assert_eq!(run_js(src), vec!["42|false"]);
}

#[test]
fn test_static_super_method_call_with_modified_arguments() {
    let src = r#"
class Base {
    static greet(name) {
        return "Hello " + name;
    }
}
class Child extends Base {
    static greet(name) {
        return super.greet(name.toUpperCase());
    }
}
console.log(Child.greet("world"));
"#;
    assert_eq!(run_js(src), vec!["Hello WORLD"]);
}

