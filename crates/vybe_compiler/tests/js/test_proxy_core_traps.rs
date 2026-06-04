/// Proxy fundamental traps — get, set, has, apply — core patterns
use super::helpers::run_js;

#[test]
fn proxy_get_intercepts_property_read() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    get(target, prop) {
        return prop in target ? target[prop] : `[${prop} not found]`;
    }
};
const obj = new Proxy({ a: 1 }, handler);
console.log(obj.a);
console.log(obj.b);
"#
        ),
        vec!["1", "[b not found]"]
    );
}

#[test]
fn proxy_set_intercepts_write() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const handler = {
    set(target, prop, value) {
        log.push(`${prop}=${value}`);
        target[prop] = value;
        return true;
    }
};
const obj = new Proxy({}, handler);
obj.x = 1;
obj.y = 2;
console.log(log.join(","));
console.log(obj.x + obj.y);
"#
        ),
        vec!["x=1,y=2", "3"]
    );
}

#[test]
fn proxy_has_intercepts_in_operator() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    has(target, prop) {
        if (prop === "secret") return false;
        return prop in target;
    }
};
const obj = new Proxy({ secret: 42, public: 1 }, handler);
console.log("public" in obj);
console.log("secret" in obj);
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn proxy_apply_wraps_function() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    apply(target, thisArg, args) {
        return target(...args) * 2;
    }
};
const double = new Proxy((x) => x + 1, handler);
console.log(double(5)); // (5+1)*2 = 12
"#
        ),
        vec!["12"]
    );
}

#[test]
fn proxy_construct_wraps_new() {
    assert_eq!(
        run_js(
            r#"
class Point { constructor(x, y) { this.x = x; this.y = y; } }
const handler = {
    construct(target, args) {
        const instance = new target(...args);
        instance.created = true;
        return instance;
    }
};
const ProxiedPoint = new Proxy(Point, handler);
const p = new ProxiedPoint(1, 2);
console.log(p.x);
console.log(p.created);
"#
        ),
        vec!["1", "true"]
    );
}

#[test]
fn proxy_revocable_after_revoke() {
    assert_eq!(
        run_js(
            r#"
const { proxy, revoke } = Proxy.revocable({ x: 1 }, {});
console.log(proxy.x);
revoke();
let threw = false;
try { proxy.x; } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["1", "true"]
    );
}

#[test]
fn proxy_read_only_property() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    set(target, prop, value) {
        if (prop === "immutable") {
            throw new TypeError("Cannot set immutable");
        }
        target[prop] = value;
        return true;
    }
};
const obj = new Proxy({ immutable: 42, mutable: 0 }, handler);
let threw = false;
try { obj.immutable = 99; } catch { threw = true; }
console.log(threw);
obj.mutable = 10;
console.log(obj.mutable);
"#
        ),
        vec!["true", "10"]
    );
}

#[test]
fn proxy_transparent_passthrough() {
    assert_eq!(
        run_js(
            r#"
const target = { x: 1, y: 2 };
const proxy = new Proxy(target, {});
proxy.z = 3;
console.log(proxy.x);
console.log(proxy.z);
console.log(target.z); // writes go to target
"#
        ),
        vec!["1", "3", "3"]
    );
}

#[test]
fn proxy_delete_property_trap() {
    assert_eq!(
        run_js(
            r#"
const deleted = [];
const handler = {
    deleteProperty(target, prop) {
        deleted.push(prop);
        return delete target[prop];
    }
};
const obj = new Proxy({ a: 1, b: 2 }, handler);
delete obj.a;
console.log(deleted.join(","));
console.log("a" in obj);
"#
        ),
        vec!["a", "false"]
    );
}

#[test]
fn proxy_own_keys_filters() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    ownKeys(target) {
        return Object.keys(target).filter(k => !k.startsWith("_"));
    }
};
const obj = new Proxy({ a: 1, _private: 2, b: 3 }, handler);
// Reflect.ownKeys must include all keys for proxy to work,
// but getOwnPropertyDescriptor will be called for each key returned
// For simplicity test that the filtered ownKeys works for Object.keys
// by also providing getOwnPropertyDescriptor
const handler2 = {
    ...handler,
    getOwnPropertyDescriptor(target, prop) {
        return Object.getOwnPropertyDescriptor(target, prop);
    }
};
const obj2 = new Proxy({ a: 1, _private: 2, b: 3 }, handler2);
const keys = Object.keys(obj2);
console.log(keys.join(","));
"#
        ),
        vec!["a,b"]
    );
}
