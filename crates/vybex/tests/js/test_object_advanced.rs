/// JavaScript Object static methods and prototype chain:
/// Object.freeze, Object.seal, Object.create, Object.is,
/// Object.defineProperty, Object.getPrototypeOf, Object.getOwnPropertyNames,
/// property descriptors, prototype manipulation.

use super::helpers::run_js;

// ===================================================================
// OBJECT.FREEZE
// ===================================================================

#[test] fn object_freeze_basic() {
    assert_eq!(run_js(r#"
let obj = { x: 1, y: 2 };
Object.freeze(obj);
obj.x = 99;
obj.z = 3;
console.log(obj.x);
console.log(obj.z);
"#), &["1", "undefined"]);
}

#[test] fn object_is_frozen() {
    assert_eq!(run_js(r#"
let obj = { a: 1 };
console.log(Object.isFrozen(obj));
Object.freeze(obj);
console.log(Object.isFrozen(obj));
"#), &["false", "true"]);
}

// ===================================================================
// OBJECT.SEAL
// ===================================================================

#[test] fn object_seal_basic() {
    assert_eq!(run_js(r#"
let obj = { x: 1, y: 2 };
Object.seal(obj);
obj.x = 99;
obj.z = 3;
console.log(obj.x);
console.log(obj.z);
"#), &["99", "undefined"]);
}

#[test] fn object_is_sealed() {
    assert_eq!(run_js(r#"
let obj = { a: 1 };
console.log(Object.isSealed(obj));
Object.seal(obj);
console.log(Object.isSealed(obj));
"#), &["false", "true"]);
}

// ===================================================================
// OBJECT.CREATE
// ===================================================================

#[test] fn object_create_basic() {
    assert_eq!(run_js(r#"
let proto = {
    greet() { return "Hello, " + this.name; }
};
let obj = Object.create(proto);
obj.name = "Alice";
console.log(obj.greet());
"#), &["Hello, Alice"]);
}

#[test] fn object_create_null_no_proto() {
    assert_eq!(run_js(r#"
let obj = Object.create(null);
obj.x = 42;
console.log(obj.x);
console.log(obj.toString);
"#), &["42", "undefined"]);
}

#[test] fn object_create_chain() {
    assert_eq!(run_js(r#"
let animal = { type: "animal", speak() { return "..."; } };
let dog = Object.create(animal);
dog.type = "dog";
dog.speak = function() { return "woof"; };
let puppy = Object.create(dog);
puppy.type = "puppy";
console.log(puppy.speak());
console.log(puppy.type);
"#), &["woof", "puppy"]);
}

// ===================================================================
// OBJECT.IS
// ===================================================================

#[test] fn object_is_comparison() {
    assert_eq!(run_js(r#"
console.log(Object.is(42, 42));
console.log(Object.is("foo", "foo"));
console.log(Object.is(NaN, NaN));
console.log(Object.is(0, -0));
console.log(Object.is(null, undefined));
"#), &["true", "true", "true", "false", "false"]);
}

// ===================================================================
// OBJECT.DEFINEPROPERTY
// ===================================================================

#[test] fn define_property_getter_setter() {
    assert_eq!(run_js(r#"
let obj = { _name: "Alice" };
Object.defineProperty(obj, "name", {
    get() { return this._name.toUpperCase(); },
    set(val) { this._name = val; }
});
console.log(obj.name);
obj.name = "Bob";
console.log(obj.name);
"#), &["ALICE", "BOB"]);
}

#[test] fn define_property_non_writable() {
    assert_eq!(run_js(r#"
let obj = {};
Object.defineProperty(obj, "PI", {
    value: 3.14159,
    writable: false,
    enumerable: true
});
console.log(obj.PI);
obj.PI = 0;
console.log(obj.PI);
"#), &["3.14159", "3.14159"]);
}

#[test] fn define_property_non_enumerable() {
    assert_eq!(run_js(r#"
let obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", {
    value: "secret",
    enumerable: false
});
console.log(Object.keys(obj).join(","));
console.log(obj.hidden);
"#), &["a,b", "secret"]);
}

// ===================================================================
// OBJECT.GETOWNPROPERTYNAMES
// ===================================================================

#[test] fn get_own_property_names() {
    assert_eq!(run_js(r#"
let obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
console.log(Object.keys(obj).join(","));
console.log(Object.getOwnPropertyNames(obj).join(","));
"#), &["a,b", "a,b,hidden"]);
}

// ===================================================================
// OBJECT.GETPROTOTYPEOF
// ===================================================================

#[test] fn get_prototype_of() {
    assert_eq!(run_js(r#"
class Animal {}
class Dog extends Animal {}
let d = new Dog();
console.log(Object.getPrototypeOf(d) === Dog.prototype);
console.log(Object.getPrototypeOf(Dog.prototype) === Animal.prototype);
"#), &["true", "true"]);
}

// ===================================================================
// PROPERTY ENUMERATION
// ===================================================================

#[test] fn for_in_own_vs_inherited() {
    assert_eq!(run_js(r#"
let parent = { a: 1 };
let child = Object.create(parent);
child.b = 2;
let own = [];
let all = [];
for (let key in child) {
    all.push(key);
    if (child.hasOwnProperty(key)) own.push(key);
}
console.log(own.join(","));
console.log(all.sort().join(","));
"#), &["b", "a,b"]);
}

#[test] fn object_keys_no_inherited() {
    assert_eq!(run_js(r#"
let parent = { x: 1 };
let child = Object.create(parent);
child.y = 2;
child.z = 3;
console.log(Object.keys(child).join(","));
"#), &["y,z"]);
}

// ===================================================================
// PROPERTY SHORTHAND AND COMPUTED
// ===================================================================

#[test] fn method_shorthand_in_object() {
    assert_eq!(run_js(r#"
let calc = {
    add(a, b) { return a + b; },
    mul(a, b) { return a * b; }
};
console.log(calc.add(3, 4));
console.log(calc.mul(3, 4));
"#), &["7", "12"]);
}

#[test] fn computed_keys_dynamic() {
    assert_eq!(run_js(r#"
let field = "name";
let obj = { [field]: "Alice", [field + "Length"]: 5 };
console.log(obj.name);
console.log(obj.nameLength);
"#), &["Alice", "5"]);
}

// ===================================================================
// OBJECT.ASSIGN DEEP PATTERNS
// ===================================================================

#[test] fn object_assign_multiple_sources() {
    assert_eq!(run_js(r#"
let a = { x: 1 };
let b = { y: 2 };
let c = { z: 3 };
let merged = Object.assign({}, a, b, c);
console.log(merged.x + "," + merged.y + "," + merged.z);
"#), &["1,2,3"]);
}

#[test] fn object_assign_override_order() {
    assert_eq!(run_js(r#"
let defaults = { color: "red", size: 10, bold: false };
let user = { color: "blue", bold: true };
let result = Object.assign({}, defaults, user);
console.log(result.color);
console.log(result.size);
console.log(result.bold);
"#), &["blue", "10", "true"]);
}

#[test] fn get_own_property_descriptor_value_flags() {
    assert_eq!(run_js(r#"
let obj = { a: 1 };
let d = Object.getOwnPropertyDescriptor(obj, "a");
console.log(d.value);
console.log(d.writable);
console.log(d.enumerable);
console.log(d.configurable);
"#), &["1", "true", "true", "true"]);
}

#[test] fn object_defineproperties_multiple_fields() {
    assert_eq!(run_js(r#"
let obj = {};
Object.defineProperties(obj, {
    a: { value: 1, enumerable: true },
    b: { value: 2, enumerable: true }
});
console.log(obj.a);
console.log(obj.b);
console.log(Object.keys(obj).join(","));
"#), &["1", "2", "a,b"]);
}

#[test] fn object_prevent_extensions_blocks_new_properties() {
    assert_eq!(run_js(r#"
let obj = { a: 1 };
Object.preventExtensions(obj);
obj.b = 2;
console.log(Object.isExtensible(obj));
console.log(obj.b);
"#), &["false", "undefined"]);
}

#[test] fn object_create_with_property_descriptors() {
    assert_eq!(run_js(r#"
let obj = Object.create({}, {
    x: { value: 5, enumerable: true },
    y: { value: 7, enumerable: true }
});
console.log(obj.x + obj.y);
console.log(Object.keys(obj).join(","));
"#), &["12", "x,y"]);
}

#[test] fn object_getprototypeof_object_create_chain() {
    assert_eq!(run_js(r#"
let proto = { kind: "base" };
let obj = Object.create(proto);
console.log(Object.getPrototypeOf(obj) === proto);
console.log(obj.kind);
"#), &["true", "base"]);
}
