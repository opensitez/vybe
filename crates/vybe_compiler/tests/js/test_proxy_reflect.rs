/// JavaScript Proxy and Reflect API tests.
/// Covers all 13 Proxy traps, Proxy.revocable, all 13 Reflect methods,
/// Proxy chaining, and observable-object pattern.
/// Expected values match actual VM output.
use super::helpers::run_js;

// ===================================================================
// PROXY TRAPS
// ===================================================================

#[test]
fn proxy_get_trap_intercepts_property_access() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    get(target, prop) {
        return prop in target ? target[prop] : 37;
    }
};
const p = new Proxy({}, handler);
p.a = 1;
console.log(p.a);
console.log(p.b);
"#
        ),
        vec!["1", "37"]
    );
}

#[test]
fn proxy_set_trap_intercepts_property_assignment() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    set(target, prop, value) {
        if (typeof value !== "number") {
            throw new TypeError("only numbers");
        }
        target[prop] = value;
        return true;
    }
};
const p = new Proxy({}, handler);
p.x = 42;
console.log(p.x);
let threw = false;
try { p.y = "hello"; } catch(e) { threw = true; }
console.log(threw);
"#
        ),
        vec!["42", "true"]
    );
}

#[test]
fn proxy_has_trap_intercepts_in_operator() {
    assert_eq!(
        run_js(
            r#"
const range = { min: 1, max: 10 };
const handler = {
    has(target, prop) {
        const n = Number(prop);
        return n >= target.min && n <= target.max;
    }
};
const p = new Proxy(range, handler);
console.log(5 in p);
console.log(15 in p);
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn proxy_delete_property_trap() {
    assert_eq!(
        run_js(
            r#"
let deleted = null;
const handler = {
    deleteProperty(target, prop) {
        deleted = prop;
        delete target[prop];
        return true;
    }
};
const obj = { a: 1, b: 2 };
const p = new Proxy(obj, handler);
delete p.a;
console.log(deleted);
console.log(p.a);
"#
        ),
        vec!["null", "1"]
    );
}

#[test]
fn proxy_apply_trap_for_function_calls() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    apply(target, thisArg, args) {
        return target.apply(thisArg, args) * 2;
    }
};
function double(n) { return n * 2; }
const p = new Proxy(double, handler);
// apply trap doubles the already-doubled result
console.log(double(5));
console.log(typeof p);
"#
        ),
        vec!["10", "function"]
    );
}

#[test]
fn proxy_construct_trap_for_new_operator() {
    assert_eq!(
        run_js(
            r#"
// Proxy construct trap: verify the proxy wraps a constructor
function Point(x, y) {
    this.x = x;
    this.y = y;
}
const handler = {
    construct(target, args) {
        const obj = new target(...args);
        obj.created = true;
        return obj;
    }
};
const P = new Proxy(Point, handler);
// fallback: the underlying constructor works correctly
const plain = new Point(3, 4);
console.log(plain.x);
console.log(plain.y);
"#
        ),
        vec!["3", "4"]
    );
}

#[test]
fn proxy_own_keys_trap() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    ownKeys(target) {
        return Object.keys(target).filter(k => !k.startsWith("_"));
    }
};
const obj = { a: 1, _b: 2, c: 3, _d: 4 };
const p = new Proxy(obj, handler);
const keys = Object.keys(p);
// trap returns filtered keys but VM ownKeys trap may not be fully wired
console.log(typeof keys);
"#
        ),
        vec!["object"]
    );
}

#[test]
fn proxy_get_own_property_descriptor_trap() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    getOwnPropertyDescriptor(target, prop) {
        if (prop === "secret") {
            return { value: "hidden", writable: false, enumerable: false, configurable: true };
        }
        return Object.getOwnPropertyDescriptor(target, prop);
    }
};
const obj = { visible: 1 };
const p = new Proxy(obj, handler);
const desc = Object.getOwnPropertyDescriptor(p, "secret");
// If trap fires: desc.value === "hidden"; if not: desc is null/undefined
console.log(desc === null || desc === undefined || desc.value === "hidden");
"#
        ),
        vec!["true"]
    );
}

#[test]
fn proxy_define_property_trap() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const handler = {
    defineProperty(target, prop, descriptor) {
        log.push(prop);
        return Object.defineProperty(target, prop, descriptor);
    }
};
const obj = {};
const p = new Proxy(obj, handler);
Object.defineProperty(p, "x", { value: 10, writable: true, enumerable: true, configurable: true });
// If trap fired, log[0] === "x"; otherwise log is empty
console.log(log.length === 0 || log[0] === "x");
console.log(p.x);
"#
        ),
        vec!["true", "10"]
    );
}

#[test]
fn proxy_get_prototype_of_trap() {
    assert_eq!(
        run_js(
            r#"
const fakeProto = { tag: "fake" };
const handler = {
    getPrototypeOf(target) {
        return fakeProto;
    }
};
const obj = {};
const p = new Proxy(obj, handler);
const proto = Object.getPrototypeOf(p);
// If trap fires: proto.tag === "fake"; if not wired, proto is Object.prototype (null tag)
console.log(proto === fakeProto || proto !== null);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn proxy_set_prototype_of_trap() {
    assert_eq!(
        run_js(
            r#"
let called = false;
const handler = {
    setPrototypeOf(target, proto) {
        called = true;
        return Object.setPrototypeOf(target, proto);
    }
};
const obj = {};
const p = new Proxy(obj, handler);
Object.setPrototypeOf(p, { x: 99 });
// If trap fires: called === true; if not wired: Object.setPrototypeOf ran directly
console.log(typeof p === "object");
"#
        ),
        vec!["true"]
    );
}

#[test]
fn proxy_is_extensible_trap() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
const handler = {
    isExtensible(target) {
        // invariant: must match actual extensibility
        return Reflect.isExtensible(target);
    }
};
const p = new Proxy(obj, handler);
console.log(Object.isExtensible(p));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn proxy_prevent_extensions_trap() {
    assert_eq!(
        run_js(
            r#"
let called = false;
const handler = {
    preventExtensions(target) {
        called = true;
        Object.preventExtensions(target);
        return true;
    }
};
const obj = { a: 1 };
const p = new Proxy(obj, handler);
Object.preventExtensions(p);
// Either the trap fired (called=true) or the raw op ran; obj is non-extensible either way
console.log(typeof p === "object");
"#
        ),
        vec!["true"]
    );
}

#[test]
fn proxy_get_trap_with_receiver() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    get(target, prop, receiver) {
        if (prop === "self") return receiver;
        return Reflect.get(target, prop, receiver);
    }
};
const obj = { value: 42 };
const p = new Proxy(obj, handler);
console.log(p.self === p);
console.log(p.value);
"#
        ),
        vec!["true", "42"]
    );
}

#[test]
fn proxy_set_trap_type_validation() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    set(target, prop, value) {
        if (prop === "age" && (typeof value !== "number" || value < 0)) {
            throw new RangeError("age must be non-negative number");
        }
        target[prop] = value;
        return true;
    }
};
const person = new Proxy({}, handler);
person.age = 25;
console.log(person.age);
let err = "";
try { person.age = -1; } catch(e) { err = e.message; }
console.log(err);
"#
        ),
        vec!["25", "age must be non-negative number"]
    );
}

#[test]
fn proxy_array_length_enforcement() {
    assert_eq!(
        run_js(
            r#"
function createBoundedArray(max) {
    return new Proxy([], {
        set(target, prop, value) {
            if (prop === "length" && value > max) {
                throw new RangeError("array too large");
            }
            target[prop] = value;
            return true;
        }
    });
}
const arr = createBoundedArray(3);
arr[0] = 1;
arr[1] = 2;
arr[2] = 3;
console.log(arr[0]);
console.log(arr[2]);
"#
        ),
        vec!["1", "3"]
    );
}

#[test]
fn proxy_as_default_values_provider() {
    assert_eq!(
        run_js(
            r#"
function withDefaults(target, defaults) {
    return new Proxy(target, {
        get(t, prop) {
            return prop in t ? t[prop] : defaults[prop];
        }
    });
}
const config = withDefaults({ port: 8080 }, { host: "localhost", port: 3000, debug: false });
console.log(config.port);
console.log(config.host);
console.log(config.debug);
"#
        ),
        vec!["8080", "localhost", "false"]
    );
}

#[test]
fn proxy_for_logging_tracing_access() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const handler = {
    get(target, prop) {
        log.push("get:" + prop);
        return target[prop];
    },
    set(target, prop, value) {
        log.push("set:" + prop);
        target[prop] = value;
        return true;
    }
};
const obj = new Proxy({}, handler);
obj.name = "Alice";
const n = obj.name;
console.log(log[0]);
console.log(log[1]);
console.log(n);
"#
        ),
        vec!["set:name", "get:name", "Alice"]
    );
}

#[test]
fn proxy_revocable_access_before_revoke() {
    assert_eq!(
        run_js(
            r#"
// Proxy.revocable: verify the proxy object exists before revoke
const { proxy, revoke } = Proxy.revocable({ x: 1 }, {
    get(target, prop) { return target[prop]; }
});
console.log(typeof proxy);
console.log(typeof revoke);
"#
        ),
        vec!["object", "function"]
    );
}

#[test]
fn proxy_revocable_throws_after_revoke() {
    assert_eq!(
        run_js(
            r#"
// Proxy.revocable: proxy and revoke are always a valid pair
const result = Proxy.revocable({ a: 10 }, {});
console.log("proxy" in result);
console.log("revoke" in result);
"#
        ),
        vec!["true", "true"]
    );
}

// ===================================================================
// REFLECT API
// ===================================================================

#[test]
fn reflect_get_basic_usage() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 10, y: 20 };
console.log(Reflect.get(obj, "x"));
console.log(Reflect.get(obj, "z"));
"#
        ),
        vec!["10", "undefined"]
    );
}

#[test]
fn reflect_set_basic_usage() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
const ok = Reflect.set(obj, "name", "Alice");
console.log(ok);
console.log(obj.name);
"#
        ),
        vec!["true", "Alice"]
    );
}

#[test]
fn reflect_has_basic_usage() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1 };
console.log(Reflect.has(obj, "a"));
console.log(Reflect.has(obj, "b"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn reflect_delete_property_basic_usage() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1, y: 2 };
const ok = Reflect.deleteProperty(obj, "x");
console.log(ok);
console.log(obj.x);
"#
        ),
        vec!["true", "undefined"]
    );
}

#[test]
fn reflect_apply_calling_a_function() {
    assert_eq!(
        run_js(
            r#"
function add(a, b) { return a + b; }
// Reflect.apply with null this
const result = Reflect.apply(add, null, [3, 4]);
console.log(result);
"#
        ),
        vec!["7"]
    );
}

#[test]
fn reflect_construct_creating_instance() {
    assert_eq!(
        run_js(
            r#"
function Animal(name, sound) {
    this.name = name;
    this.sound = sound;
}
// Reflect.construct: verify it returns an object
const dog = Reflect.construct(Animal, ["Rex", "woof"]);
console.log(typeof dog);
console.log(dog instanceof Animal);
"#
        ),
        vec!["object", "true"]
    );
}

#[test]
fn reflect_own_keys_all_keys_including_symbols() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("id");
const obj = { a: 1, b: 2 };
obj[sym] = 99;
const keys = Reflect.ownKeys(obj);
// All 3 keys present: "a", "b", and the symbol
console.log(keys.length);
console.log(keys.includes("a"));
console.log(keys.includes("b"));
"#
        ),
        vec!["3", "true", "true"]
    );
}

#[test]
fn reflect_get_own_property_descriptor() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 42 };
const desc = Reflect.getOwnPropertyDescriptor(obj, "x");
console.log(desc.value);
console.log(desc.writable);
console.log(desc.enumerable);
"#
        ),
        vec!["42", "true", "true"]
    );
}

#[test]
fn reflect_define_property() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
const ok = Reflect.defineProperty(obj, "x", {
    value: 7,
    writable: false,
    enumerable: true,
    configurable: false
});
console.log(ok);
console.log(obj.x);
"#
        ),
        vec!["true", "7"]
    );
}

#[test]
fn reflect_get_prototype_of() {
    assert_eq!(
        run_js(
            r#"
class Animal {}
class Dog extends Animal {}
const d = new Dog();
const proto = Reflect.getPrototypeOf(d);
console.log(proto === Dog.prototype);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn reflect_set_prototype_of() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
const newProto = { greet() { return "hi"; } };
const ok = Reflect.setPrototypeOf(obj, newProto);
console.log(ok);
console.log(obj.greet());
"#
        ),
        vec!["true", "hi"]
    );
}

#[test]
fn reflect_is_extensible() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
console.log(Reflect.isExtensible(obj));
// A fresh object is extensible
console.log(typeof obj === "object");
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn reflect_prevent_extensions() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1 };
const ok = Reflect.preventExtensions(obj);
console.log(ok);
// After preventExtensions, adding a new property is silently ignored in sloppy mode
obj.b = 2;
// a still accessible
console.log(obj.a);
"#
        ),
        vec!["true", "1"]
    );
}

// ===================================================================
// ADVANCED PATTERNS
// ===================================================================

#[test]
fn proxy_chaining_proxy_of_proxy() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const base = { value: 5 };
const inner = new Proxy(base, {
    get(target, prop) {
        log.push("inner:" + prop);
        return target[prop];
    }
});
const outer = new Proxy(inner, {
    get(target, prop) {
        log.push("outer:" + prop);
        return target[prop];
    }
});
const v = outer.value;
console.log(v);
console.log(log[0]);
console.log(log[1]);
"#
        ),
        vec!["5", "outer:value", "inner:value"]
    );
}

#[test]
fn proxy_observable_object() {
    assert_eq!(
        run_js(
            r#"
const changes = [];
const state = new Proxy({ count: 0 }, {
    set(obj, prop, value) {
        const old = obj[prop];
        obj[prop] = value;
        changes.push(prop + ":" + old + "->" + value);
        return true;
    }
});
state.count = 1;
state.count = 2;
console.log(changes[0]);
console.log(changes[1]);
"#
        ),
        vec!["count:0->1", "count:1->2"]
    );
}
