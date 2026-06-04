/// Prototype-based patterns — inheritance, delegation, mixins
use super::helpers::run_js;

#[test]
fn classical_inheritance_via_prototype() {
    assert_eq!(
        run_js(
            r#"
function Animal(name) { this.name = name; }
Animal.prototype.speak = function() { return this.name + " speaks"; };
function Dog(name) { Animal.call(this, name); }
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.bark = function() { return this.name + " barks"; };
const d = new Dog("Rex");
console.log(d.speak());
console.log(d.bark());
console.log(d instanceof Dog);
console.log(d instanceof Animal);
"#
        ),
        vec!["Rex speaks", "Rex barks", "true", "true"]
    );
}

#[test]
fn parasitic_inheritance() {
    assert_eq!(
        run_js(
            r#"
function createEnhanced(original) {
    const clone = Object.create(original);
    clone.describe = function() { return "Enhanced: " + this.name; };
    return clone;
}
const base = { name: "Base", greet() { return "Hello from " + this.name; } };
const enhanced = createEnhanced(base);
enhanced.name = "Enhanced";
console.log(enhanced.greet());
console.log(enhanced.describe());
console.log(Object.getPrototypeOf(enhanced) === base);
"#
        ),
        vec!["Hello from Enhanced", "Enhanced: Enhanced", "true"]
    );
}

#[test]
fn mixin_composition() {
    assert_eq!(
        run_js(
            r#"
const Serializable = (superclass) => class extends superclass {
    serialize() { return JSON.stringify(this); }
};
const Timestamped = (superclass) => class extends superclass {
    constructor(...args) { super(...args); this.createdAt = 0; }
};
class Base {
    constructor(name) { this.name = name; }
}
class User extends Timestamped(Serializable(Base)) {}
const u = new User("Alice");
const s = u.serialize();
console.log(JSON.parse(s).name);
console.log("createdAt" in u);
console.log(u instanceof Base);
"#
        ),
        vec!["Alice", "true", "true"]
    );
}

#[test]
fn property_delegation() {
    assert_eq!(
        run_js(
            r#"
function delegate(target, host, methods) {
    for (const m of methods) {
        host[m] = (...args) => target[m](...args);
    }
    return host;
}
class Stack {
    #arr = [];
    push(v) { this.#arr.push(v); return this; }
    pop() { return this.#arr.pop(); }
    peek() { return this.#arr[this.#arr.length - 1]; }
    get size() { return this.#arr.length; }
}
const queue = delegate(new Stack(), {}, ["push", "pop", "peek"]);
queue.push(1); queue.push(2);
console.log(queue.peek());
console.log(queue.pop());
"#
        ),
        vec!["2", "2"]
    );
}

#[test]
fn prototype_augmentation_chain() {
    assert_eq!(
        run_js(
            r#"
function addMethods(proto, methods) {
    Object.assign(proto, methods);
    return proto;
}
const base = { greet() { return "Hi"; } };
const extended = Object.create(addMethods(base, {
    goodbye() { return "Bye"; }
}));
extended.name = "Alice";
console.log(extended.greet());
console.log(extended.goodbye());
console.log(Object.getPrototypeOf(extended) === base);
"#
        ),
        vec!["Hi", "Bye", "true"]
    );
}

#[test]
fn method_borrowing() {
    assert_eq!(
        run_js(
            r#"
const arrayLike = { 0: "a", 1: "b", 2: "c", length: 3 };
const joined = Array.prototype.join.call(arrayLike, "-");
const mapped = Array.prototype.map.call(arrayLike, s => s.toUpperCase());
console.log(joined);
console.log(mapped.join(","));
"#
        ),
        vec!["a-b-c", "A,B,C"]
    );
}

#[test]
fn symbol_iterator_on_prototype() {
    assert_eq!(
        run_js(
            r#"
function Range(start, end) { this.start = start; this.end = end; }
Range.prototype[Symbol.iterator] = function*() {
    for (let i = this.start; i <= this.end; i++) yield i;
};
const r = new Range(1, 5);
console.log([...r].join(","));
console.log(Array.from(r).join(","));
"#
        ),
        vec!["1,2,3,4,5", "1,2,3,4,5"]
    );
}

#[test]
fn null_prototype_safe_dict() {
    assert_eq!(
        run_js(
            r#"
const dict = Object.create(null);
dict.constructor = "fake";
dict.toString = "fake";
dict.hasOwnProperty = "fake";
// Null prototype means no inherited methods
console.log(Object.getPrototypeOf(dict));
console.log("constructor" in dict);
// Object.hasOwn works on null-prototype
dict.real = 42;
console.log(Object.hasOwn(dict, "real"));
"#
        ),
        vec!["null", "true", "true"]
    );
}
