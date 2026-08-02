// vybe-test: js/module_import_patterns/conditional_module_loading_simulation
// origin: languages/js/tests/js/test_module_import_patterns.rs

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

async function loadModule(name) {
    if (name === "a") return { value: 1 };
    if (name === "b") return { value: 2 };
    throw new Error("unknown:" + name);
}

async function main() {
    const mod = await loadModule("a");
    console.log(mod.value);
    try {
        await loadModule("c");
    } catch (e) {
        console.log(e.message);
    }
}
main();
