/// Miscellaneous JS patterns — computed properties, getters/setters, Proxy advanced
use super::helpers::run_js;

#[test]
fn computed_property_keys_advanced() {
    assert_eq!(
        run_js(
            r#"
const prefix = "get";
const methods = ["Name", "Age", "Email"].reduce((obj, field) => {
    obj[`${prefix}${field}`] = () => `getting ${field}`;
    return obj;
}, {});
console.log(methods.getName());
console.log(methods.getAge());
"#
        ),
        vec!["getting Name", "getting Age"]
    );
}

#[test]
fn getter_setter_validation() {
    assert_eq!(
        run_js(
            r#"
class Temperature {
    #celsius = 0;
    get celsius() { return this.#celsius; }
    set celsius(v) {
        if (v < -273.15) throw new RangeError("Below absolute zero");
        this.#celsius = v;
    }
    get fahrenheit() { return this.#celsius * 9/5 + 32; }
    set fahrenheit(v) { this.celsius = (v - 32) * 5/9; }
}
const t = new Temperature();
t.celsius = 100;
console.log(t.fahrenheit);
t.fahrenheit = 32;
console.log(t.celsius);
let threw = false;
try { t.celsius = -300; } catch { threw = true; }
console.log(threw);
"#
        ),
        vec!["212", "0", "true"]
    );
}

#[test]
fn proxy_logging_trap() {
    assert_eq!(
        run_js(
            r#"
function logged(obj) {
    const log = [];
    const proxy = new Proxy(obj, {
        get(target, prop, receiver) {
            log.push("get:" + String(prop));
            return Reflect.get(target, prop, receiver);
        },
        set(target, prop, value, receiver) {
            log.push("set:" + String(prop));
            return Reflect.set(target, prop, value, receiver);
        }
    });
    return [proxy, log];
}
const [p, log] = logged({ x: 1 });
p.x;
p.y = 2;
p.x;
console.log(log.join(","));
"#
        ),
        vec!["get:x,set:y,get:x"]
    );
}

#[test]
fn reflect_vs_direct_access() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 1 };
Object.defineProperty(obj, "y", { value: 2, configurable: false, writable: false, enumerable: true });
console.log(Reflect.get(obj, "x"));
console.log(Reflect.has(obj, "y"));
console.log(Reflect.ownKeys(obj).join(","));
console.log(Reflect.deleteProperty(obj, "x"));
console.log(Reflect.deleteProperty(obj, "y"));
"#
        ),
        vec!["1", "true", "x,y", "true", "false"]
    );
}

#[test]
fn object_create_descriptors() {
    assert_eq!(
        run_js(
            r#"
const proto = {
    greet() { return `Hello, ${this.name}`; }
};
const obj = Object.create(proto, {
    name: { value: "Alice", writable: true, enumerable: true, configurable: true },
    age: { value: 30, writable: false, enumerable: true, configurable: false }
});
console.log(obj.greet());
obj.name = "Bob";
console.log(obj.name);
obj.age = 99;
console.log(obj.age);
"#
        ),
        vec!["Hello, Alice", "Bob", "30"]
    );
}

#[test]
fn property_access_order() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperty(obj, "computed", {
    get() { return this._x * 2; },
    configurable: true
});
obj._x = 5;
console.log(obj.computed);
Object.defineProperty(obj, "computed", {
    get() { return this._x * 3; }
});
console.log(obj.computed);
"#
        ),
        vec!["10", "15"]
    );
}

#[test]
fn argument_object_vs_rest() {
    assert_eq!(
        run_js(
            r#"
function withArgs() {
    return [...arguments].map(x => x * 2);
}
function withRest(...args) {
    return args.map(x => x * 2);
}
console.log(withArgs(1, 2, 3).join(","));
console.log(withRest(1, 2, 3).join(","));
console.log(Array.isArray(withRest()));
"#
        ),
        vec!["2,4,6", "2,4,6", "true"]
    );
}

#[test]
fn object_seal_vs_freeze_behavior() {
    assert_eq!(
        run_js(
            r#"
const sealed = Object.seal({ x: 1, y: 2 });
sealed.x = 99;
sealed.z = 3;
delete sealed.x;
console.log(sealed.x);
console.log("z" in sealed);
console.log(Object.isSealed(sealed));

const frozen = Object.freeze({ a: 1 });
frozen.a = 99;
console.log(frozen.a);
console.log(Object.isFrozen(frozen));
"#
        ),
        vec!["99", "false", "true", "1", "true"]
    );
}

#[test]
fn prototype_lookup_chain() {
    assert_eq!(
        run_js(
            r#"
const base = { type: "base", describe() { return this.type + ":" + this.name; } };
const mid = Object.create(base);
mid.type = "mid";
const leaf = Object.create(mid);
leaf.name = "leaf";
console.log(leaf.describe());
console.log(leaf.type);
console.log(leaf.hasOwnProperty("name"));
console.log(leaf.hasOwnProperty("type"));
"#
        ),
        vec!["mid:leaf", "mid", "true", "false"]
    );
}

#[test]
fn property_enumeration_full() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("s");
const obj = Object.create({ inherited: true });
obj.own = 1;
Object.defineProperty(obj, "nonEnum", { value: 2, enumerable: false });
obj[sym] = "symbol";
console.log(Object.keys(obj).join(","));
console.log(Object.getOwnPropertyNames(obj).sort().join(","));
console.log(Reflect.ownKeys(obj).filter(k => typeof k === "string").sort().join(","));
"#
        ),
        vec!["own", "nonEnum,own", "nonEnum,own"]
    );
}
