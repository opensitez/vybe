// vybe-test: js/class_private_advanced/private_field_serialization_tojson
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class Person {
    #name;
    #age;
    constructor(name, age) { this.#name = name; this.#age = age; }
    toJSON() { return { name: this.#name, age: this.#age }; }
    serialize() { return JSON.stringify(this.toJSON()); }
}
const p = new Person("Alice", 30);
const json = p.serialize();
__check(__line(json), "{\"name\":\"Alice\",\"age\":30}");
