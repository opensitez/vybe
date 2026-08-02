// vybe-test: js/module_patterns/plugin_system_pattern
// origin: languages/js/tests/js/test_module_patterns.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

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
__check(__line(system.run("double", 21)), "42");
__check(__line(system.run("greet", "World")), "Hello, World!");
