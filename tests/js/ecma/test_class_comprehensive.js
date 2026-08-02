// vybe-test: js/ecma/test_class_comprehensive
// origin: languages/js/tests/js/js_ecma_test.rs

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
            constructor(make, year) {
                this.make = make;
                this.year = year;
                this.speed = 0;
            }
            accelerate(amount) {
                this.speed = this.speed + amount;
            }
            brake(amount) {
                this.speed = Math.max(0, this.speed - amount);
            }
            describe() {
                return `${this.year} ${this.make} going ${this.speed}mph`;
            }
        }
        
        let car = new Vehicle("Tesla", 2024);
        car.accelerate(60);
        car.accelerate(20);
        car.brake(10);
        __check(__line(car.describe()), "2024 Tesla going 70mph");
