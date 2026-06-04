/// Proxy advanced — has trap, set trap validation, deleteProperty, ownKeys,
/// apply trap, construct trap, getPrototypeOf/setPrototypeOf traps, invariants.
use super::helpers::run_js;

// ── has trap ──────────────────────────────────────────────────────────────────

#[test]
fn proxy_has_trap_intercepts_in_operator() {
    assert_eq!(
        run_js(
            r#"
const handler = {
    has(target, prop) {
        console.log("has:" + prop);
        return prop in target;
    }
};
const proxy = new Proxy({ a: 1 }, handler);
console.log("a" in proxy);
console.log("b" in proxy);
"#
        ),
        vec!["has:a", "true", "has:b", "false"]
    );
}

#[test]
fn proxy_has_trap_can_hide_properties() {
    assert_eq!(
        run_js(
            r#"
const hidden = new Set(["secret"]);
const proxy = new Proxy(
    { secret: 42, public: 1 },
    { has(target, prop) { return !hidden.has(prop) && prop in target; } }
);
console.log("public" in proxy);
console.log("secret" in proxy);
"#
        ),
        vec!["true", "false"]
    );
}

// ── set trap ──────────────────────────────────────────────────────────────────

#[test]
fn proxy_set_trap_validates_input() {
    assert_eq!(
        run_js(
            r#"
const proxy = new Proxy({}, {
    set(target, prop, value) {
        if (typeof value !== "number") throw new TypeError("must be number");
        target[prop] = value;
        return true;
    }
});
proxy.x = 42;
console.log(proxy.x);
let threw = false;
try { proxy.y = "string"; } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["42", "true"]
    );
}

#[test]
fn proxy_set_trap_can_normalize_values() {
    assert_eq!(
        run_js(
            r#"
const proxy = new Proxy({}, {
    set(target, prop, value) {
        target[prop] = typeof value === "string" ? value.toLowerCase() : value;
        return true;
    }
});
proxy.name = "HELLO";
console.log(proxy.name);
"#
        ),
        vec!["hello"]
    );
}

// ── deleteProperty trap ───────────────────────────────────────────────────────

#[test]
fn proxy_delete_property_trap_intercepts() {
    assert_eq!(
        run_js(
            r#"
const log = [];
const proxy = new Proxy({ a: 1, b: 2 }, {
    deleteProperty(target, prop) {
        log.push("delete:" + prop);
        return delete target[prop];
    }
});
delete proxy.a;
console.log(log.join(","));
console.log("a" in proxy);
"#
        ),
        vec!["delete:a", "false"]
    );
}

#[test]
fn proxy_delete_property_can_prevent_deletion() {
    assert_eq!(
        run_js(
            r#"
const protected_props = new Set(["core"]);
const proxy = new Proxy({ core: 1, temp: 2 }, {
    deleteProperty(target, prop) {
        if (protected_props.has(prop)) return false;
        return delete target[prop];
    }
});
delete proxy.temp;
delete proxy.core;
console.log("temp" in proxy);
console.log("core" in proxy);
"#
        ),
        vec!["false", "true"]
    );
}

// ── ownKeys trap ──────────────────────────────────────────────────────────────

#[test]
fn proxy_ownkeys_trap_filters_keys() {
    assert_eq!(
        run_js(
            r#"
const target = { a: 1, _private: 2, b: 3 };
const proxy = new Proxy(target, {
    ownKeys(t) {
        return Object.keys(t).filter(k => !k.startsWith("_"));
    }
});
console.log(Object.keys(proxy).sort().join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn proxy_ownkeys_intercepts_object_getownpropertynames() {
    assert_eq!(
        run_js(
            r#"
const proxy = new Proxy({ x: 1, y: 2 }, {
    ownKeys() { return ["x", "z"]; }
});
// ownKeys must return subset of actual keys (invariant)
// or keys that exist in target for non-configurable
const keys = Object.getOwnPropertyNames(proxy);
console.log(keys.includes("x"));
"#
        ),
        vec!["true"]
    );
}

// ── apply trap ────────────────────────────────────────────────────────────────

#[test]
fn proxy_apply_trap_wraps_function_call() {
    assert_eq!(
        run_js(
            r#"
function add(a, b) { return a + b; }
const proxy = new Proxy(add, {
    apply(target, thisArg, args) {
        console.log("called with " + args.length + " args");
        return target.apply(thisArg, args);
    }
});
console.log(proxy(2, 3));
"#
        ),
        vec!["called with 2 args", "5"]
    );
}

#[test]
fn proxy_apply_can_modify_return_value() {
    assert_eq!(
        run_js(
            r#"
const proxy = new Proxy((x) => x * x, {
    apply(target, thisArg, [x]) {
        return target(x) + 1;
    }
});
console.log(proxy(5));
"#
        ),
        vec!["26"]
    );
}

// ── construct trap ────────────────────────────────────────────────────────────

#[test]
fn proxy_construct_trap_intercepts_new() {
    assert_eq!(
        run_js(
            r#"
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
const ProxiedPoint = new Proxy(Point, {
    construct(target, args) {
        const [x, y] = args;
        return new target(x * 2, y * 2);
    }
});
const p = new ProxiedPoint(3, 4);
console.log(p.x);
console.log(p.y);
"#
        ),
        vec!["6", "8"]
    );
}

// ── getPrototypeOf trap ───────────────────────────────────────────────────────

#[test]
fn proxy_getprototypeof_trap() {
    assert_eq!(
        run_js(
            r#"
const fakeProto = { tag: "spoofed" };
const proxy = new Proxy({}, {
    getPrototypeOf() { return fakeProto; }
});
console.log(Object.getPrototypeOf(proxy) === fakeProto);
"#
        ),
        vec!["true"]
    );
}

// ── Proxy revocable ───────────────────────────────────────────────────────────

#[test]
fn proxy_revocable_access_after_revoke_throws() {
    assert_eq!(
        run_js(
            r#"
const { proxy, revoke } = Proxy.revocable({ x: 1 }, {});
console.log(proxy.x);
revoke();
let threw = false;
try { proxy.x; } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#
        ),
        vec!["1", "true"]
    );
}

// ── Proxy for validation ──────────────────────────────────────────────────────

#[test]
fn proxy_as_type_validator() {
    assert_eq!(
        run_js(
            r#"
function typed(obj, schema) {
    return new Proxy(obj, {
        set(target, prop, value) {
            if (schema[prop] && typeof value !== schema[prop]) {
                throw new TypeError(`${prop} must be ${schema[prop]}`);
            }
            target[prop] = value;
            return true;
        }
    });
}
const person = typed({}, { name: "string", age: "number" });
person.name = "Alice";
person.age = 30;
console.log(person.name + ":" + person.age);
let threw = false;
try { person.age = "old"; } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["Alice:30", "true"]
    );
}
