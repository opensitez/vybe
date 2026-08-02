// vybe-test: js/module_patterns/factory_function_pattern
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
__check(__line(alice.hasPermission("delete")), "true");
__check(__line(bob.hasPermission("delete")), "false");
__check(__line(String(alice)), "Alice (admin)");
