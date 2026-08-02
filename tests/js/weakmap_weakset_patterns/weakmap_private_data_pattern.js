// vybe-test: js/weakmap_weakset_patterns/weakmap_private_data_pattern
// origin: languages/js/tests/js/test_weakmap_weakset_patterns.rs

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

const privateData = new WeakMap();
class Person {
    constructor(name, age) {
        privateData.set(this, { name, age });
    }
    greet() {
        const { name } = privateData.get(this);
        return "Hi, I'm " + name;
    }
    get age() { return privateData.get(this).age; }
}
const p = new Person("Alice", 30);
__check(__line(p.greet()), "Hi, I'm Alice");
__check(__line(p.age), "30");
__check(__line(p.name), "undefined"); // undefined — not on instance
