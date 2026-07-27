/// JavaScript class patterns: mixins, static blocks, accessor patterns,
/// abstract-like patterns, builder pattern, singleton, observer,
/// class expressions, toString/valueOf, Symbol.toPrimitive, Symbol.hasInstance.
use super::helpers::run_js;

// ===================================================================
// MIXINS
// ===================================================================

#[test]
fn mixin_pattern() {
    assert_eq!(
        run_js(
            r#"
let Serializable = (Base) => class extends Base {
    serialize() { return JSON.stringify(this); }
};
let Loggable = (Base) => class extends Base {
    log() { console.log("LOG: " + this.name); }
};
class User {
    constructor(name) { this.name = name; }
}
class EnhancedUser extends Loggable(Serializable(User)) {}
let u = new EnhancedUser("Alice");
u.log();
let s = u.serialize();
console.log(s.includes("Alice"));
"#
        ),
        &["LOG: Alice", "true"]
    );
}

#[test]
fn mixin_multiple_methods() {
    assert_eq!(
        run_js(
            r#"
function Timestamped(Base) {
    return class extends Base {
        getTimestamp() { return "2024-01-01"; }
    };
}
function Tagged(Base) {
    return class extends Base {
        setTag(tag) { this._tag = tag; }
        getTag() { return this._tag; }
    };
}
class Item {}
class TaggedItem extends Tagged(Timestamped(Item)) {}
let item = new TaggedItem();
item.setTag("important");
console.log(item.getTag());
console.log(item.getTimestamp());
"#
        ),
        &["important", "2024-01-01"]
    );
}

// ===================================================================
// STATIC BLOCKS
// ===================================================================

#[test]
fn static_block_initialization() {
    assert_eq!(
        run_js(
            r#"
class Config {
    static values;
    static {
        Config.values = { debug: false, version: "1.0" };
    }
}
console.log(Config.values.version);
console.log(Config.values.debug);
"#
        ),
        &["1.0", "false"]
    );
}

#[test]
fn static_block_computed() {
    assert_eq!(
        run_js(
            r#"
class MathConstants {
    static PI;
    static TAU;
    static {
        MathConstants.PI = 3.14159;
        MathConstants.TAU = MathConstants.PI * 2;
    }
}
console.log(MathConstants.PI);
console.log(MathConstants.TAU);
"#
        ),
        &["3.14159", "6.28318"]
    );
}

// ===================================================================
// BUILDER PATTERN
// ===================================================================

#[test]
fn builder_pattern() {
    assert_eq!(
        run_js(
            r#"
class QueryBuilder {
    constructor() { this.parts = []; }
    select(fields) { this.parts.push("SELECT " + fields); return this; }
    from(table) { this.parts.push("FROM " + table); return this; }
    where(cond) { this.parts.push("WHERE " + cond); return this; }
    build() { return this.parts.join(" "); }
}
let q = new QueryBuilder()
    .select("*")
    .from("users")
    .where("age > 18")
    .build();
console.log(q);
"#
        ),
        &["SELECT * FROM users WHERE age > 18"]
    );
}

// ===================================================================
// SINGLETON PATTERN
// ===================================================================

#[test]
fn singleton_pattern() {
    assert_eq!(
        run_js(
            r#"
class Database {
    static instance = null;
    constructor(name) { this.name = name; }
    static getInstance(name) {
        if (!Database.instance) {
            Database.instance = new Database(name);
        }
        return Database.instance;
    }
}
let db1 = Database.getInstance("main");
let db2 = Database.getInstance("other");
console.log(db1 === db2);
console.log(db2.name);
"#
        ),
        &["true", "main"]
    );
}

// ===================================================================
// OBSERVER PATTERN
// ===================================================================

#[test]
fn observer_pattern() {
    assert_eq!(
        run_js(
            r#"
class EventEmitter {
    constructor() { this.listeners = {}; }
    on(event, fn) {
        if (!this.listeners[event]) this.listeners[event] = [];
        this.listeners[event].push(fn);
    }
    emit(event, ...args) {
        if (this.listeners[event]) {
            this.listeners[event].forEach(fn => fn(...args));
        }
    }
}
let emitter = new EventEmitter();
emitter.on("data", val => console.log("got: " + val));
emitter.on("data", val => console.log("also: " + val));
emitter.emit("data", 42);
"#
        ),
        &["got: 42", "also: 42"]
    );
}

// ===================================================================
// TOSTRING / VALUEOF
// ===================================================================

#[test]
fn tostring_override() {
    assert_eq!(
        run_js(
            r#"
class Money {
    constructor(amount, currency) {
        this.amount = amount;
        this.currency = currency;
    }
    toString() { return this.amount + " " + this.currency; }
}
let m = new Money(100, "USD");
console.log("" + m);
console.log(`${m}`);
"#
        ),
        &["100 USD", "100 USD"]
    );
}

#[test]
fn valueof_override() {
    assert_eq!(
        run_js(
            r#"
class Num {
    constructor(v) { this.v = v; }
    valueOf() { return this.v; }
}
let a = new Num(10);
let b = new Num(20);
console.log(a + b);
console.log(a * 3);
"#
        ),
        &["30", "30"]
    );
}

// ===================================================================
// SYMBOL.TOPRIMITIVE
// ===================================================================

#[test]
fn symbol_toprimitive() {
    assert_eq!(
        run_js(
            r#"
class Temperature {
    constructor(celsius) { this.celsius = celsius; }
    [Symbol.toPrimitive](hint) {
        if (hint === "number") return this.celsius;
        if (hint === "string") return this.celsius + "°C";
        return this.celsius;
    }
}
let t = new Temperature(100);
console.log(+t);
console.log(`${t}`);
"#
        ),
        &["100", "100°C"]
    );
}

// ===================================================================
// SYMBOL.HASINSTANCE
// ===================================================================

#[test]
fn symbol_hasinstance() {
    assert_eq!(
        run_js(
            r#"
class Even {
    static [Symbol.hasInstance](num) {
        return typeof num === "number" && num % 2 === 0;
    }
}
console.log(4 instanceof Even);
console.log(3 instanceof Even);
"#
        ),
        &["true", "false"]
    );
}

// ===================================================================
// CLASS EXPRESSION
// ===================================================================

#[test]
fn class_expression_anonymous() {
    assert_eq!(
        run_js(
            r#"
let Animal = class {
    constructor(name) { this.name = name; }
    speak() { return this.name + " speaks"; }
};
let a = new Animal("Cat");
console.log(a.speak());
"#
        ),
        &["Cat speaks"]
    );
}

#[test]
fn class_expression_named() {
    assert_eq!(
        run_js(
            r#"
let Foo = class Bar {
    static className() { return "Bar"; }
};
console.log(Foo.className());
"#
        ),
        &["Bar"]
    );
}

// ===================================================================
// GETTER/SETTER PATTERNS
// ===================================================================

#[test]
fn computed_getter_setter() {
    assert_eq!(
        run_js(
            r#"
class Circle {
    constructor(radius) { this._radius = radius; }
    get radius() { return this._radius; }
    set radius(r) {
        if (r < 0) throw new Error("negative");
        this._radius = r;
    }
    get area() { return 3.14 * this._radius * this._radius; }
    get circumference() { return 2 * 3.14 * this._radius; }
}
let c = new Circle(5);
console.log(c.area);
console.log(c.circumference);
c.radius = 10;
console.log(c.area);
"#
        ),
        &["78.5", "31.400000000000002", "314"]
    );
}

#[test]
fn getter_setter_validation() {
    assert_eq!(
        run_js(
            r#"
class User {
    #email;
    get email() { return this.#email; }
    set email(val) {
        if (!val.includes("@")) throw new Error("invalid email");
        this.#email = val;
    }
}
let u = new User();
u.email = "alice@test.com";
console.log(u.email);
try {
    u.email = "invalid";
} catch (e) {
    console.log(e.message);
}
"#
        ),
        &["alice@test.com", "invalid email"]
    );
}

// ===================================================================
// ABSTRACT-LIKE PATTERN
// ===================================================================

#[test]
fn abstract_class_pattern() {
    assert_eq!(
        run_js(
            r#"
class Shape {
    area() { throw new Error("not implemented"); }
}
class Rect extends Shape {
    constructor(w, h) { super(); this.w = w; this.h = h; }
    area() { return this.w * this.h; }
}
class Circle extends Shape {
    constructor(r) { super(); this.r = r; }
    area() { return 3.14 * this.r * this.r; }
}
let shapes = [new Rect(3, 4), new Circle(5)];
shapes.forEach(s => console.log(s.area()));
try {
    new Shape().area();
} catch (e) {
    console.log(e.message);
}
"#
        ),
        &["12", "78.5", "not implemented"]
    );
}

// ===================================================================
// ITERABLE CLASS
// ===================================================================

#[test]
fn class_iterable_with_symbol_iterator() {
    assert_eq!(
        run_js(
            r#"
class NumberRange {
    constructor(start, end) { this.start = start; this.end = end; }
    [Symbol.iterator]() {
        let current = this.start;
        let end = this.end;
        return {
            next() {
                if (current <= end) return { value: current++, done: false };
                return { done: true };
            }
        };
    }
}
let nums = [...new NumberRange(1, 5)];
console.log(nums.join(","));
"#
        ),
        &["1,2,3,4,5"]
    );
}

#[test]
fn class_method_extracted_loses_this_binding() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    constructor() { this.value = 3; }
    get() { return this && this.value; }
}
let c = new Counter();
let fn = c.get;
console.log(c.get());
console.log(fn());
"#
        ),
        &["3", "undefined"]
    );
}

#[test]
fn static_field_is_shared_on_class_not_instance() {
    assert_eq!(
        run_js(
            r#"
class Box {
    static count = 2;
}
let b = new Box();
console.log(Box.count);
console.log(b.count);
"#
        ),
        &["2", "undefined"]
    );
}

#[test]
fn subclass_can_call_super_method() {
    assert_eq!(
        run_js(
            r#"
class Animal {
    speak() { return "animal"; }
}
class Dog extends Animal {
    speak() { return super.speak() + " dog"; }
}
console.log(new Dog().speak());
"#
        ),
        &["animal dog"]
    );
}

#[test]
fn class_instance_fields_are_per_instance() {
    assert_eq!(
        run_js(
            r#"
class Counter {
    value = 0;
    inc() { this.value += 1; }
}
let a = new Counter();
let b = new Counter();
a.inc();
a.inc();
b.inc();
console.log(a.value);
console.log(b.value);
"#
        ),
        &["2", "1"]
    );
}

#[test]
fn class_constructor_can_return_custom_object() {
    assert_eq!(
        run_js(
            r#"
class Weird {
    constructor() {
        this.value = 1;
        return { value: 99 };
    }
}
let w = new Weird();
console.log(w.value);
"#
        ),
        &["99"]
    );
}

#[test]
fn class_extends_expression() {
    assert_eq!(
        run_js(
            r#"
function makeBase() {
    return class {
        greet() { return "hi"; }
    };
}
class Derived extends makeBase() {}
console.log(new Derived().greet());
"#
        ),
        &["hi"]
    );
}

#[test]
fn getter_runs_on_property_access_each_time() {
    assert_eq!(
        run_js(
            r#"
class Seq {
    constructor() { this.n = 0; }
    get next() { this.n += 1; return this.n; }
}
let s = new Seq();
console.log(s.next);
console.log(s.next);
"#
        ),
        &["1", "2"]
    );
}

#[test]
fn setter_can_normalize_input() {
    assert_eq!(
        run_js(
            r#"
class User {
    set name(value) { this._name = value.trim(); }
    get name() { return this._name; }
}
let u = new User();
u.name = "  Alice  ";
console.log(u.name);
"#
        ),
        &["Alice"]
    );
}

#[test]
fn symbol_to_primitive_default_hint_used_in_addition() {
    assert_eq!(
        run_js(
            r#"
class Amount {
    constructor(v) { this.v = v; }
    [Symbol.toPrimitive](hint) {
        console.log(hint);
        return this.v;
    }
}
let a = new Amount(7);
console.log(a + 5);
"#
        ),
        &["default", "12"]
    );
}

// ===================================================================
// USER AUGMENTATION OF INTRINSIC PROTOTYPES
// ===================================================================
//
// Assigning a user function to a PRIMITIVE's intrinsic prototype and reaching
// it from a primitive receiver. Distinct from the cases above, which only READ
// builtins off a prototype (`String.prototype.slice.call(...)`). ECMA-262
// §6.1.5: member access on a primitive boxes it and consults the intrinsic
// prototype, so a user-added member must resolve exactly like a builtin one.
//
// Arrays are covered separately and already work — they are ordinary objects,
// so they never exercised this path. See flexclassplan.md §4d.

#[test]
fn number_prototype_user_method_resolves_off_primitive() {
    assert_eq!(
        run_js(
            r#"
Number.prototype.doubled = function () { return this * 2; };
console.log((5).doubled());
"#
        ),
        vec!["10"]
    );
}

#[test]
fn string_prototype_user_method_resolves_off_primitive() {
    assert_eq!(
        run_js(
            r#"
String.prototype.shout = function () { return this + "!"; };
console.log("hi".shout());
"#
        ),
        vec!["hi!"]
    );
}

#[test]
fn array_prototype_user_method_resolves_off_literal() {
    assert_eq!(
        run_js(
            r#"
Array.prototype.second = function () { return this[1]; };
console.log([1, 2, 3].second());
"#
        ),
        vec!["2"]
    );
}

#[test]
fn intrinsic_prototype_augmentation_does_not_shadow_builtin() {
    assert_eq!(
        run_js(
            r#"
Number.prototype.doubled = function () { return this * 2; };
console.log((5).toFixed(2));
"#
        ),
        vec!["5.00"]
    );
}
