// vybe-test: js/prototype_oop_patterns/prototype_augmentation_chain
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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

function addMethods(proto, methods) {
    Object.assign(proto, methods);
    return proto;
}
const base = { greet() { return "Hi"; } };
const extended = Object.create(addMethods(base, {
    goodbye() { return "Bye"; }
}));
extended.name = "Alice";
__check(__line(extended.greet()), "Hi");
__check(__line(extended.goodbye()), "Bye");
__check(__line(Object.getPrototypeOf(extended) === base), "true");
