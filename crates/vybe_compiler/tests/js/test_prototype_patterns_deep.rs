/// Prototype and class patterns — prototype chain manipulation, delegation

use super::helpers::run_js;

#[test]
fn prototype_assignment_via_object_create() {
    assert_eq!(run_js(r#"
const animal = {
    speak() { return this.name + " says " + this.sound; }
};
const dog = Object.create(animal);
dog.name = "Rex";
dog.sound = "woof";
console.log(dog.speak());
"#), vec!["Rex says woof"]);
}

#[test]
fn prototype_chain_method_lookup() {
    assert_eq!(run_js(r#"
function Animal(name) { this.name = name; }
Animal.prototype.speak = function() { return this.name; };
function Dog(name, breed) {
    Animal.call(this, name);
    this.breed = breed;
}
Dog.prototype = Object.create(Animal.prototype);
Dog.prototype.constructor = Dog;
Dog.prototype.bark = function() { return this.name + " barks!"; };
const d = new Dog("Rex", "Lab");
console.log(d.speak());
console.log(d.bark());
console.log(d instanceof Dog);
console.log(d instanceof Animal);
"#), vec!["Rex", "Rex barks!", "true", "true"]);
}

#[test]
fn object_create_with_property_descriptors() {
    assert_eq!(run_js(r#"
const proto = { greet() { return "Hello from " + this.name; } };
const obj = Object.create(proto, {
    name: { value: "Alice", writable: true, enumerable: true, configurable: true }
});
console.log(obj.name);
console.log(obj.greet());
"#), vec!["Alice", "Hello from Alice"]);
}

#[test]
fn set_prototype_of_changes_behavior() {
    assert_eq!(run_js(r#"
const a = { whoAmI() { return "A"; } };
const b = { whoAmI() { return "B"; } };
const obj = Object.create(a);
console.log(obj.whoAmI());
Object.setPrototypeOf(obj, b);
console.log(obj.whoAmI());
"#), vec!["A", "B"]);
}

#[test]
fn get_prototype_of_traversal() {
    assert_eq!(run_js(r#"
class A {}
class B extends A {}
class C extends B {}
const c = new C();
const chain = [];
let proto = Object.getPrototypeOf(c);
while (proto !== null) {
    if (proto.constructor) chain.push(proto.constructor.name);
    proto = Object.getPrototypeOf(proto);
}
console.log(chain.includes("C"));
console.log(chain.includes("B"));
console.log(chain.includes("A"));
"#), vec!["true", "true", "true"]);
}

#[test]
fn hasownproperty_vs_in() {
    assert_eq!(run_js(r#"
const proto = { inherited: 1 };
const obj = Object.create(proto);
obj.own = 2;
console.log(obj.hasOwnProperty("own"));
console.log(obj.hasOwnProperty("inherited"));
console.log("own" in obj);
console.log("inherited" in obj);
"#), vec!["true", "false", "true", "true"]);
}

#[test]
fn object_hasown_static_method() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
const nullProto = Object.create(null);
nullProto.key = "value";
console.log(Object.hasOwn(obj, "a"));
// Object.hasOwn works even on null-prototype objects
console.log(Object.hasOwn(nullProto, "key"));
"#), vec!["true", "true"]);
}

#[test]
fn prototype_pollution_defense_pattern() {
    assert_eq!(run_js(r#"
// Using null-prototype objects as safe maps
const safe = Object.create(null);
safe.key = "value";
console.log(safe.key);
// No inherited methods
console.log(typeof safe.toString);
console.log(typeof safe.hasOwnProperty);
"#), vec!["value", "undefined", "undefined"]);
}

#[test]
fn constructor_property_preserved() {
    assert_eq!(run_js(r#"
class Foo {}
const f = new Foo();
console.log(f.constructor === Foo);
console.log(f.constructor.name);
"#), vec!["true", "Foo"]);
}
