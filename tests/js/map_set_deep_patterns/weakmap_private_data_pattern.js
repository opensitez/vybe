// vybe-test: js/map_set_deep_patterns/weakmap_private_data_pattern
// origin: languages/js/tests/js/test_map_set_deep_patterns.rs

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

const _private = new WeakMap();
class Person {
    constructor(name, age) {
        _private.set(this, { name, age });
    }
    get name() { return _private.get(this).name; }
    get age() { return _private.get(this).age; }
    birthday() { _private.get(this).age++; }
}
const p = new Person("Alice", 30);
__check(__line(p.name), "Alice");
__check(__line(p.age), "30");
p.birthday();
__check(__line(p.age), "31");
__check(__line(_private.has(p)), "true");
