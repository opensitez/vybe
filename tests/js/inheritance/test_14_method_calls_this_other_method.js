// vybe-test: js/inheritance/test_14_method_calls_this_other_method
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

class Calc {
            double(x) { return x * 2; }
            quadruple(x) { return this.double(this.double(x)); }
        }
        let c = new Calc();
        __check(__line(c.quadruple(3)), "12");
