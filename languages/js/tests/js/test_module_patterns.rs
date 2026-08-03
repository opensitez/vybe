/// Module-like patterns — revealing module, namespace, factory, singleton
use super::helpers::run_js;

#[test]
fn revealing_module_pattern() {
    assert_eq!(
        run_js(
            r#"
const BankAccount = (() => {
    let balance = 0;
    const deposit = (n) => { balance += n; };
    const withdraw = (n) => { if (n > balance) throw new Error("insufficient"); balance -= n; };
    const getBalance = () => balance;
    return { deposit, withdraw, getBalance };
})();
BankAccount.deposit(100);
BankAccount.deposit(50);
BankAccount.withdraw(30);
console.log(BankAccount.getBalance());
"#
        ),
        vec!["120"]
    );
}

#[test]
fn namespace_pattern() {
    assert_eq!(
        run_js(
            r#"
const App = {
    utils: {
        add: (a, b) => a + b,
        multiply: (a, b) => a * b },
    config: {
        version: "1.0",
        debug: false },
    init() {
        return `App v${this.config.version} initialized`;
    }
};
console.log(App.utils.add(3, 4));
console.log(App.init());
console.log(App.config.debug);
"#
        ),
        vec!["7", "App v1.0 initialized", "false"]
    );
}

#[test]
fn factory_function_pattern() {
    assert_eq!(
        run_js(
            r#"
function createUser(name, role = "user") {
    const permissions = role === "admin" ? ["read", "write", "delete"] : ["read"];
    return {
        name,
        role,
        hasPermission(p) { return permissions.includes(p); },
        toString() { return `${name} (${role})`; }
    };
}
const alice = createUser("Alice", "admin");
const bob = createUser("Bob");
console.log(alice.hasPermission("delete"));
console.log(bob.hasPermission("delete"));
console.log(String(alice));
"#
        ),
        vec!["true", "false", "Alice (admin)"]
    );
}

#[test]
fn plugin_system_pattern() {
    assert_eq!(
        run_js(
            r#"
class PluginSystem {
    #plugins = new Map();
    register(name, plugin) { this.#plugins.set(name, plugin); }
    run(name, ...args) {
        const plugin = this.#plugins.get(name);
        if (!plugin) throw new Error(`Plugin ${name} not found`);
        return plugin(...args);
    }
}
const system = new PluginSystem();
system.register("double", x => x * 2);
system.register("greet", name => `Hello, ${name}!`);
console.log(system.run("double", 21));
console.log(system.run("greet", "World"));
"#
        ),
        vec!["42", "Hello, World!"]
    );
}

#[test]
fn dependency_injection_pattern() {
    assert_eq!(
        run_js(
            r#"
class Logger {
    log(msg) { return "[LOG] " + msg; }
}
class UserService {
    constructor(logger) { this.logger = logger; }
    createUser(name) {
        const msg = this.logger.log(`Creating user: ${name}`);
        return msg;
    }
}
const logger = new Logger();
const service = new UserService(logger);
console.log(service.createUser("Alice"));
"#
        ),
        vec!["[LOG] Creating user: Alice"]
    );
}

#[test]
fn module_with_state() {
    assert_eq!(
        run_js(
            r#"
const Counter = (() => {
    let count = 0;
    return {
        inc: () => ++count,
        dec: () => --count,
        reset: () => { count = 0; return count; },
        value: () => count };
})();
Counter.inc();
Counter.inc();
Counter.inc();
Counter.dec();
console.log(Counter.value());
Counter.reset();
console.log(Counter.value());
"#
        ),
        vec!["2", "0"]
    );
}

#[test]
fn lazy_initialization_pattern() {
    assert_eq!(
        run_js(
            r#"
class LazyInit {
    #instance = null;
    #factory;
    constructor(factory) { this.#factory = factory; }
    get() {
        if (!this.#instance) {
            this.#instance = this.#factory();
        }
        return this.#instance;
    }
}
let created = 0;
const lazy = new LazyInit(() => { created++; return { value: 42 }; });
console.log(created);     // not yet created
const v = lazy.get();
console.log(created);     // now created
const v2 = lazy.get();
console.log(created);     // not created again
console.log(v === v2);    // same instance
"#
        ),
        vec!["0", "1", "1", "true"]
    );
}

#[test]
fn memoize_with_max_size() {
    assert_eq!(
        run_js(
            r#"
function lruMemoize(fn, maxSize = 3) {
    const cache = new Map();
    return function(key) {
        if (cache.has(key)) {
            const val = cache.get(key);
            cache.delete(key);
            cache.set(key, val); // move to end (most recent)
            return val;
        }
        const result = fn(key);
        if (cache.size >= maxSize) {
            cache.delete(cache.keys().next().value); // remove oldest
        }
        cache.set(key, result);
        return result;
    };
}
let calls = 0;
const sq = lruMemoize(x => { calls++; return x * x; }, 2);
sq(2); sq(3); sq(2); sq(4); // sq(4) evicts sq(3)
sq(3); // must recompute (evicted)
console.log(calls); // 2+3+4+3 computed = 4 unique + 1 recompute = 5
"#
        ),
        vec!["5"]
    );
}
