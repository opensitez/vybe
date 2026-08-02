// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_construct_basic_instantiation
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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
    constructor(name, age) {
        this.name = name;
        this.age = age;
    }
}
const p = Reflect.construct(Person, ["Alice", 30]);
__check(__line(`${p.name}:${p.age}|isPerson=${p instanceof Person}`), "Alice:30|isPerson=true");
