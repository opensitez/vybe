/// WeakMap and WeakSet patterns — expando, brand checking, private data
use super::helpers::run_js;

#[test]
fn weakmap_stores_per_object_data() {
    assert_eq!(
        run_js(
            r#"
const data = new WeakMap();
const obj1 = {};
const obj2 = {};
data.set(obj1, { id: 1 });
data.set(obj2, { id: 2 });
console.log(data.get(obj1).id);
console.log(data.get(obj2).id);
console.log(data.has(obj1));
"#
        ),
        vec!["1", "2", "true"]
    );
}

#[test]
fn weakmap_delete_removes_entry() {
    assert_eq!(
        run_js(
            r#"
const wm = new WeakMap();
const key = {};
wm.set(key, 42);
console.log(wm.has(key));
wm.delete(key);
console.log(wm.has(key));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn weakmap_requires_object_key() {
    assert_eq!(
        run_js(
            r#"
const wm = new WeakMap();
let threw = false;
try { wm.set("string", 1); } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn weakmap_private_data_pattern() {
    assert_eq!(
        run_js(
            r#"
const privateData = new WeakMap();
class Person {
    constructor(name, age) {
        privateData.set(this, { name, age });
    }
    greet() {
        const { name } = privateData.get(this);
        return "Hi, I'm " + name;
    }
    get age() { return privateData.get(this).age; }
}
const p = new Person("Alice", 30);
console.log(p.greet());
console.log(p.age);
console.log(p.name); // undefined — not on instance
"#
        ),
        vec!["Hi, I'm Alice", "30", "undefined"]
    );
}

#[test]
fn weakset_tracks_objects() {
    assert_eq!(
        run_js(
            r#"
const seen = new WeakSet();
const a = {}, b = {}, c = {};
seen.add(a);
seen.add(b);
console.log(seen.has(a));
console.log(seen.has(c));
seen.delete(a);
console.log(seen.has(a));
"#
        ),
        vec!["true", "false", "false"]
    );
}

#[test]
fn weakset_brand_checking() {
    assert_eq!(
        run_js(
            r#"
const validInstances = new WeakSet();
class Secure {
    constructor() {
        validInstances.add(this);
    }
    static validate(obj) {
        if (!validInstances.has(obj)) throw new TypeError("Invalid instance");
        return true;
    }
}
const s = new Secure();
console.log(Secure.validate(s));
let threw = false;
try { Secure.validate({}); } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn weakset_requires_object_values() {
    assert_eq!(
        run_js(
            r#"
const ws = new WeakSet();
let threw = false;
try { ws.add(42); } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn weakmap_as_cache() {
    assert_eq!(
        run_js(
            r#"
const cache = new WeakMap();
function process(obj) {
    if (cache.has(obj)) return cache.get(obj);
    const result = Object.keys(obj).length;
    cache.set(obj, result);
    return result;
}
const o = { a: 1, b: 2, c: 3 };
console.log(process(o));
console.log(process(o)); // from cache
"#
        ),
        vec!["3", "3"]
    );
}
