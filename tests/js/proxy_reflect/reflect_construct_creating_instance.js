// vybe-test: js/proxy_reflect/reflect_construct_creating_instance
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

function Animal(name, sound) {
    this.name = name;
    this.sound = sound;
}
// Reflect.construct: verify it returns an object
const dog = Reflect.construct(Animal, ["Rex", "woof"]);
__check(__line(typeof dog), "object");
__check(__line(dog instanceof Animal), "true");
