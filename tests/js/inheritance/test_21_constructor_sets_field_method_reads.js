// vybe-test: js/inheritance/test_21_constructor_sets_field_method_reads
// origin: languages/js/tests/js/js_inheritance_test.rs

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
            info() { return this.name + " is " + this.age; }
        }
        let p = new Person("Alice", 30);
        __check(__line(p.info()), "Alice is 30");
