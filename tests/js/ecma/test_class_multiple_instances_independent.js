// vybe-test: js/ecma/test_class_multiple_instances_independent
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

class Box {
            constructor(w, h) {
                this.w = w;
                this.h = h;
            }
            area() { return this.w * this.h; }
        }
        let a = new Box(3, 4);
        let b = new Box(5, 6);
        __check(__line(a.area(), b.area()), "12 30");
