/// Getter/setter deep — property descriptor, inheritance, enumerable/configurable,
/// computed accessor names, accessor in mixin, getter side effects.
use super::helpers::run_js;

// ── basic getter/setter ───────────────────────────────────────────────────────

#[test]
fn getter_called_on_every_access() {
    assert_eq!(
        run_js(
            r#"
let count = 0;
const obj = {
    get value() { return ++count; }
};
obj.value; obj.value;
console.log(count);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn setter_intercepts_assignment() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    _x: 0,
    get x() { return this._x; },
    set x(v) { this._x = v < 0 ? 0 : v; }
};
obj.x = 5;
console.log(obj.x);
obj.x = -3;
console.log(obj.x);
"#
        ),
        vec!["5", "0"]
    );
}

// ── getter in prototype ───────────────────────────────────────────────────────

#[test]
fn getter_on_prototype_accessed_by_instances() {
    assert_eq!(
        run_js(
            r#"
function Foo(x) { this.x = x; }
Object.defineProperty(Foo.prototype, "doubled", {
    get() { return this.x * 2; }
});
const a = new Foo(5);
const b = new Foo(10);
console.log(a.doubled);
console.log(b.doubled);
"#
        ),
        vec!["10", "20"]
    );
}

// ── getter not enumerable by default in class ─────────────────────────────────

#[test]
fn class_getter_not_in_object_keys() {
    assert_eq!(
        run_js(
            r#"
class C {
    get name() { return "test"; }
}
const c = new C();
const keys = Object.keys(c);
console.log(keys.includes("name"));
// class getters are non-enumerable
"#
        ),
        vec!["false"]
    );
}

// ── lazy initialization via getter ────────────────────────────────────────────

#[test]
fn lazy_initialization_getter_pattern() {
    assert_eq!(
        run_js(
            r#"
let computed = 0;
const obj = {
    get expensive() {
        // Replace with own property after first access
        const value = ++computed;
        Object.defineProperty(this, "expensive", { value, writable: true });
        return value;
    }
};
console.log(obj.expensive); // 1
console.log(obj.expensive); // 1 (cached — own property now)
console.log(computed);       // 1
"#
        ),
        vec!["1", "1", "1"]
    );
}

// ── computed accessor names ────────────────────────────────────────────────────

#[test]
fn computed_getter_name() {
    assert_eq!(
        run_js(
            r#"
const prop = "value";
const obj = {
    _v: 42,
    get [prop]() { return this._v; },
    set [prop](v) { this._v = v; }
};
console.log(obj.value);
obj.value = 100;
console.log(obj.value);
"#
        ),
        vec!["42", "100"]
    );
}

// ── setter-only / getter-only ──────────────────────────────────────────────────

#[test]
fn write_only_setter_reads_as_undefined() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    _log: [],
    set entry(v) { this._log.push(v); }
};
obj.entry = "a";
obj.entry = "b";
console.log(obj.entry); // undefined — no getter
console.log(obj._log.join(","));
"#
        ),
        vec!["undefined", "a,b"]
    );
}

#[test]
fn read_only_getter() {
    assert_eq!(
        run_js(
            r#"
const obj = {
    get pi() { return 3.14159; }
};
obj.pi = 99; // silently ignored in non-strict
console.log(obj.pi);
"#
        ),
        vec!["3.14159"]
    );
}

// ── defineProperty accessor ───────────────────────────────────────────────────

#[test]
#[allow(non_snake_case)]
fn defineProperty_creates_accessor() {
    assert_eq!(
        run_js(
            r#"
const obj = { _n: 0 };
Object.defineProperty(obj, "n", {
    get() { return this._n; },
    set(v) { if (Number.isInteger(v)) this._n = v; },
    enumerable: true,
    configurable: true
});
obj.n = 5;
console.log(obj.n);
obj.n = 2.5; // ignored — not integer
console.log(obj.n);
"#
        ),
        vec!["5", "5"]
    );
}

// ── inheritance of accessor ───────────────────────────────────────────────────

#[test]
fn accessor_inherited_through_class_chain() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    constructor(n) { this._name = n; }
    get name() { return this._name.toUpperCase(); }
}
class Dog extends Animal {}

const d = new Dog("rex");
console.log(d.name);
"#
        ),
        vec!["REX"]
    );
}

#[test]
fn subclass_can_override_accessor() {
    assert_eq!(
        run_js(
            r#"
class Base {
    get label() { return "Base"; }
}
class Child extends Base {
    get label() { return "Child:" + super.label; }
}
console.log(new Child().label);
"#
        ),
        vec!["Child:Base"]
    );
}

// ── getter in mixin ───────────────────────────────────────────────────────────

#[test]
fn mixin_adds_accessor_to_class() {
    assert_eq!(
        run_js(
            r#"
const Timestamped = (Base) => class extends Base {
    get timestamp() { return this._ts || 0; }
    set timestamp(v) { this._ts = v; }
};

class Record {}
class TimedRecord extends Timestamped(Record) {}

const r = new TimedRecord();
r.timestamp = 12345;
console.log(r.timestamp);
"#
        ),
        vec!["12345"]
    );
}

// ── Object.getOwnPropertyDescriptor for accessor ─────────────────────────────

#[test]
fn accessor_descriptor_does_not_have_value() {
    assert_eq!(
        run_js(
            r#"
const obj = { get x() { return 1; } };
const desc = Object.getOwnPropertyDescriptor(obj, "x");
console.log("value" in desc);
console.log(typeof desc.get);
console.log(typeof desc.set);
console.log(desc.enumerable);
"#
        ),
        vec!["false", "function", "undefined", "true"]
    );
}
