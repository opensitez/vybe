// vybe-test: js/interop/test_d41_class_expression
// origin: languages/js/tests/js/js_interop_test.rs

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

let MyClass = class {
            constructor(x) { this.x = x; }
            get() { return this.x; }
        };
        let obj = new MyClass(42);
        __check(__line(obj.get()), "42");
