// vybe-test: js/interop/test_d34_class_constructor_fields_methods
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

class Rectangle {
            constructor(w, h) {
                this.width = w;
                this.height = h;
            }
            area() { return this.width * this.height; }
            perimeter() { return 2 * (this.width + this.height); }
        }
        let r = new Rectangle(3, 4);
        __check(__line(r.area(), r.perimeter()), "12 14");
