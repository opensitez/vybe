// vybe-test: js/inheritance/test_22_derived_and_parent_fields
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

class Vehicle {
            constructor(type) { this.type = type; }
        }
        class Car extends Vehicle {
            constructor(brand) {
                super("car");
                this.brand = brand;
            }
        }
        let c = new Car("Toyota");
        __check(__line(c.type, c.brand), "car Toyota");
