// vybe-test: js/ecma/test_super_with_args
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

class Base {
            constructor(x) {
                this.x = x;
            }
            getX() { return this.x; }
        }
        class Child extends Base {
            constructor(x, y) {
                super(x);
                this.y = y;
            }
            getY() { return this.y; }
            sum() { return this.x + this.y; }
        }
        let c = new Child(10, 20);
        __check(__line(c.getX(), c.getY(), c.sum()), "10 20 30");
