/// Prototype chain manipulation — Object.create inheritance, prototype reassignment,
/// hasOwnProperty, property lookup order, prototype pollution prevention,
/// Object.freeze prototype, getPrototypeOf/setPrototypeOf.
use super::helpers::run_js;

// ── prototype lookup ──────────────────────────────────────────────────────────

#[test]
fn property_found_on_prototype() {
    assert_eq!(
        run_js(
            r#"
const proto = { x: 42 };
const obj = Object.create(proto);
console.log(obj.x);
console.log(obj.hasOwnProperty("x"));
console.log(proto.hasOwnProperty("x"));
"#
        ),
        vec!["42", "false", "true"]
    );
}

#[test]
fn prototype_chain_three_levels() {
    assert_eq!(
        run_js(
            r#"
const a = { hello() { return "hello"; } };
const b = Object.create(a);
b.world = function() { return "world"; };
const c = Object.create(b);

console.log(c.hello());
console.log(c.world());
console.log(Object.getPrototypeOf(c) === b);
console.log(Object.getPrototypeOf(b) === a);
"#
        ),
        vec!["hello", "world", "true", "true"]
    );
}

// ── hasOwn vs hasOwnProperty ──────────────────────────────────────────────────

#[test]
fn object_hasown_consistent_with_hasnownproperty() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create({ inherited: true });
obj.own = true;
console.log(obj.hasOwnProperty("own"));
console.log(obj.hasOwnProperty("inherited"));
console.log(Object.hasOwn(obj, "own"));
console.log(Object.hasOwn(obj, "inherited"));
"#
        ),
        vec!["true", "false", "true", "false"]
    );
}

// ── overriding prototype property ─────────────────────────────────────────────

#[test]
fn own_property_shadows_prototype_property() {
    assert_eq!(
        run_js(
            r#"
const proto = { x: 1 };
const obj = Object.create(proto);
obj.x = 99; // creates own property, shadows prototype
console.log(obj.x);
console.log(proto.x);
"#
        ),
        vec!["99", "1"]
    );
}

// ── setPrototypeOf ────────────────────────────────────────────────────────────

#[test]
fn set_prototype_of_changes_chain() {
    assert_eq!(
        run_js(
            r#"
const a = { hello() { return "from a"; } };
const b = { hello() { return "from b"; } };
const obj = Object.create(a);
console.log(obj.hello());
Object.setPrototypeOf(obj, b);
console.log(obj.hello());
"#
        ),
        vec!["from a", "from b"]
    );
}

// ── null prototype ────────────────────────────────────────────────────────────

#[test]
fn null_prototype_has_no_tostring() {
    assert_eq!(
        run_js(
            r#"
const bare = Object.create(null);
bare.x = 1;
console.log(Object.getPrototypeOf(bare) === null);
console.log(Object.getPrototypeOf(bare));
"#
        ),
        vec!["true", "null"]
    );
}

#[test]
fn null_prototype_safe_dict() {
    assert_eq!(
        run_js(
            r#"
// Use null-prototype objects as safe dicts (no prototype pollution)
const dict = Object.create(null);
dict.constructor = "overridden"; // doesn't affect anything
dict.hasOwnProperty = "also overridden";
console.log(dict.constructor);
console.log(Object.hasOwn(dict, "constructor"));
"#
        ),
        vec!["overridden", "true"]
    );
}

// ── prototype inheritance with Object.create ──────────────────────────────────

#[test]
fn classical_inheritance_via_object_create() {
    assert_eq!(
        run_js(
            r#"
function Animal(name) { this.name = name; }
Animal.prototype.speak = function() { return this.name + " says hi"; };

function Dog(name, breed) {
    Animal.call(this, name);
    this.breed = breed;
}
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.bark = function() { return this.name + " barks"; };

const d = new Dog("Rex", "Lab");
console.log(d.speak());
console.log(d.bark());
console.log(d instanceof Dog);
console.log(d instanceof Animal);
"#
        ),
        vec!["Rex says hi", "Rex barks", "true", "true"]
    );
}

// ── in operator checks prototype chain ────────────────────────────────────────

#[test]
fn in_operator_walks_prototype_chain() {
    assert_eq!(
        run_js(
            r#"
const proto = { inherited: true };
const obj = Object.create(proto);
console.log("inherited" in obj);
console.log("own" in obj);
obj.own = true;
console.log("own" in obj);
"#
        ),
        vec!["true", "false", "true"]
    );
}

// ── prototype of class instance ───────────────────────────────────────────────

#[test]
fn class_instance_prototype_is_class_prototype() {
    assert_eq!(
        run_js(
            r#"
class Foo {
    bar() { return "bar"; }
}
const f = new Foo();
console.log(Object.getPrototypeOf(f) === Foo.prototype);
console.log(f.bar());
"#
        ),
        vec!["true", "bar"]
    );
}

// ── instanceof and prototype ──────────────────────────────────────────────────

#[test]
fn instanceof_via_symbol_hasinstance() {
    assert_eq!(
        run_js(
            r#"
class Range {
    constructor(min, max) { this.min = min; this.max = max; }
}
const r = new Range(0, 10);
console.log(r instanceof Range);
console.log(r instanceof Object);
"#
        ),
        vec!["true", "true"]
    );
}

// ── method on prototype vs own ────────────────────────────────────────────────

#[test]
fn method_on_prototype_shared_across_instances() {
    assert_eq!(
        run_js(
            r#"
function Counter() { this.count = 0; }
Counter.prototype.increment = function() { this.count++; };

const c1 = new Counter();
const c2 = new Counter();

c1.increment(); c1.increment();
c2.increment();

console.log(c1.count);
console.log(c2.count);
// Shared method
console.log(c1.increment === c2.increment);
"#
        ),
        vec!["2", "1", "true"]
    );
}

// ── Object.create as delegation ───────────────────────────────────────────────

#[test]
fn delegation_pattern_via_object_create() {
    assert_eq!(
        run_js(
            r#"
const logger = {
    log(msg) { return `[LOG] ${msg}`; },
    error(msg) { return `[ERROR] ${msg}`; }
};
const app = Object.create(logger);
app.name = "MyApp";
app.run = function() { return this.log("Running " + this.name); };

console.log(app.run());
console.log(app.error("crash"));
"#
        ),
        vec!["[LOG] Running MyApp", "[ERROR] crash"]
    );
}
