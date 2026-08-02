// vybe-test: js/missing_features/constructor_calls_method
// origin: languages/js/tests/js/js_missing_features_test.rs

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

class Foo {
            constructor() {
                this.value = 0;
                this.init();
            }
            init() {
                this.value = 42;
            }
        }
        let f = new Foo();
        __check(__line(f.value), "42");
