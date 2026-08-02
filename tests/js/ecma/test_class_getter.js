// vybe-test: js/ecma/test_class_getter
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

class Person {
            constructor(first, last) {
                this.first = first;
                this.last = last;
            }
            get fullName() {
                return this.first + " " + this.last;
            }
        }
        let p = new Person("John", "Doe");
        __check(__line(p.first, p.last), "John Doe");
