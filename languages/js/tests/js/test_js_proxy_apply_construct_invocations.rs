use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Proxy Invocations Traps (apply, construct)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_proxy_apply_trap_function_call() {
    let src = r#"
function sum(a, b) { return a + b; }
const proxy = new Proxy(sum, {
    apply(target, thisArg, args) {
        return target(...args) * 10;
    }
});
console.log(proxy(2, 3));
"#;
    assert_eq!(run_js(src), vec!["50"]);
}

#[test]
fn test_js_proxy_apply_trap_this_arg_inspection() {
    let src = r#"
function getGreeting() { return `Hello ${this.name}`; }
const proxy = new Proxy(getGreeting, {
    apply(target, thisArg, args) {
        return target.apply(thisArg, args).toUpperCase();
    }
});
const ctx = { name: "Alice" };
console.log(proxy.call(ctx));
"#;
    assert_eq!(run_js(src), vec!["HELLO ALICE"]);
}

#[test]
fn test_js_proxy_construct_trap_new_operator() {
    let src = r#"
function User(name) { this.name = name; }
const proxy = new Proxy(User, {
    construct(target, args, newTarget) {
        const obj = new target(...args);
        obj.created = true;
        return obj;
    }
});
const u = new proxy("Bob");
console.log(u.name + "|" + u.created);
"#;
    assert_eq!(run_js(src), vec!["Bob|true"]);
}

#[test]
fn test_js_proxy_construct_trap_must_return_object_invariant() {
    let src = r#"
function Item() {}
const proxy = new Proxy(Item, {
    construct(target, args) {
        return 42; // Construct trap MUST return an object!
    }
});
try {
    new proxy();
} catch (e) {
    console.log("Construct Non-Object Invariant Error");
}
"#;
    assert_eq!(run_js(src), vec!["Construct Non-Object Invariant Error"]);
}

#[test]
fn test_js_proxy_apply_trap_non_callable_target_throws_typeerror() {
    let src = r#"
const nonFn = { a: 1 };
try {
    const proxy = new Proxy(nonFn, {
        apply(target, thisArg, args) { return 0; }
    });
    proxy();
} catch (e) {
    console.log("Non-Callable Apply Error");
}
"#;
    assert_eq!(run_js(src), vec!["Non-Callable Apply Error"]);
}

#[test]
fn test_js_proxy_construct_trap_non_constructor_target_throws() {
    let src = r#"
const arrowFn = () => {};
try {
    const proxy = new Proxy(arrowFn, {
        construct(target, args) { return {}; }
    });
    new proxy();
} catch (e) {
    console.log("Non-Constructor Proxy Error");
}
"#;
    assert_eq!(run_js(src), vec!["Non-Constructor Proxy Error"]);
}

#[test]
fn test_js_proxy_construct_trap_new_target_subclassing() {
    let src = r#"
class Base {
    constructor() { this.base = true; }
}
let capturedNewTarget;
const proxy = new Proxy(Base, {
    construct(target, args, newTarget) {
        capturedNewTarget = newTarget;
        return Reflect.construct(target, args, newTarget);
    }
});
class Derived extends proxy {}
const instance = new Derived();
console.log(capturedNewTarget === Derived);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_apply_trap_memoization() {
    let src = r#"
const cache = new Map();
function fib(n) {
    if (n <= 1) return n;
    return fib(n - 1) + fib(n - 2);
}
const memoFib = new Proxy(fib, {
    apply(target, thisArg, args) {
        const key = args[0];
        if (cache.has(key)) return cache.get(key);
        const res = target(...args);
        cache.set(key, res);
        return res;
    }
});
console.log(memoFib(10));
console.log(cache.has(10));
"#;
    assert_eq!(run_js(src), vec!["55", "true"]);
}

#[test]
fn test_js_proxy_construct_trap_singleton_pattern() {
    let src = r#"
let instance = null;
class Service {}
const SingletonService = new Proxy(Service, {
    construct(target, args) {
        if (!instance) instance = new target(...args);
        return instance;
    }
});
const s1 = new SingletonService();
const s2 = new SingletonService();
console.log(s1 === s2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_apply_trap_arguments_logging() {
    let src = r#"
const log = [];
function add(a, b) { return a + b; }
const loggedAdd = new Proxy(add, {
    apply(target, thisArg, args) {
        log.push(args.join("+"));
        return target(...args);
    }
});
loggedAdd(1, 2);
loggedAdd(10, 20);
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["1+2,10+20"]);
}

#[test]
fn test_js_proxy_construct_trap_override_returned_instance() {
    let src = r#"
function Person(name) { this.name = name; }
const proxy = new Proxy(Person, {
    construct(target, args) {
        return { custom: true, name: args[0].toUpperCase() };
    }
});
const p = new proxy("charlie");
console.log(p.name + "|" + p.custom);
"#;
    assert_eq!(run_js(src), vec!["CHARLIE|true"]);
}

#[test]
fn test_js_proxy_apply_trap_method_chaining_interception() {
    let src = r#"
const obj = {
    val: 0,
    add(n) { this.val += n; return this; }
};
const proxy = new Proxy(obj, {
    get(target, prop, receiver) {
        const orig = target[prop];
        if (typeof orig === "function") {
            return new Proxy(orig, {
                apply(fnTarget, thisArg, args) {
                    return fnTarget.apply(target, args);
                }
            });
        }
        return orig;
    }
});
proxy.add(5).add(10);
console.log(obj.val);
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_proxy_construct_trap_validates_constructor_parameters() {
    let src = r#"
class Product {
    constructor(price) { this.price = price; }
}
const ValidatedProduct = new Proxy(Product, {
    construct(target, args) {
        if (args[0] <= 0) throw new Error("Invalid Price");
        return new target(...args);
    }
});
try {
    new ValidatedProduct(-10);
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["Invalid Price"]);
}

#[test]
fn test_js_proxy_apply_trap_with_reflect_apply() {
    let src = r#"
function multiply(x, y) { return x * y; }
const proxy = new Proxy(multiply, {
    apply(target, thisArg, args) {
        return Reflect.apply(target, thisArg, [args[0] + 1, args[1] + 1]);
    }
});
console.log(proxy(2, 3)); // (2+1) * (3+1) = 12
"#;
    assert_eq!(run_js(src), vec!["12"]);
}

#[test]
fn test_js_proxy_apply_trap_arrow_function_this_unbound() {
    let src = r#"
const arrow = () => "arrow";
const proxy = new Proxy(arrow, {
    apply(target, thisArg, args) {
        return target.apply(thisArg, args);
    }
});
console.log(proxy.call({ custom: "ctx" }));
"#;
    assert_eq!(run_js(src), vec!["arrow"]);
}

#[test]
fn test_js_proxy_construct_trap_prototype_chain_intact() {
    let src = r#"
class Widget {}
const ProxyWidget = new Proxy(Widget, {
    construct(target, args, newTarget) {
        return Reflect.construct(target, args, newTarget);
    }
});
const w = new ProxyWidget();
console.log(w instanceof Widget);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_proxy_apply_trap_bind_compatibility() {
    let src = r#"
function greet(prefix, suffix) { return `${prefix} ${this.name} ${suffix}`; }
const proxy = new Proxy(greet, {
    apply(target, thisArg, args) {
        return target.apply(thisArg, args);
    }
});
const bound = proxy.bind({ name: "World" }, "Hello");
console.log(bound("!"));
"#;
    assert_eq!(run_js(src), vec!["Hello World !"]);
}

#[test]
fn test_js_proxy_construct_trap_default_behavior_without_handler() {
    let src = r#"
class Config {
    constructor(v) { this.v = v; }
}
const ProxyConfig = new Proxy(Config, {});
const c = new ProxyConfig(42);
console.log(c.v);
"#;
    assert_eq!(run_js(src), vec!["42"]);
}

#[test]
fn test_js_proxy_apply_trap_variadic_arguments_forwarding() {
    let src = r#"
function concatAll(...strings) { return strings.join("-"); }
const proxy = new Proxy(concatAll, {
    apply(target, thisArg, args) {
        return target(...args).toUpperCase();
    }
});
console.log(proxy("a", "b", "c"));
"#;
    assert_eq!(run_js(src), vec!["A-B-C"]);
}

#[test]
fn test_js_proxy_construct_trap_class_expression_target() {
    let src = r#"
const ProxyAnonClass = new Proxy(class {
    constructor(val) { this.val = val; }
}, {
    construct(target, args) {
        return new target(args[0] * 100);
    }
});
const obj = new ProxyAnonClass(3);
console.log(obj.val);
"#;
    assert_eq!(run_js(src), vec!["300"]);
}
