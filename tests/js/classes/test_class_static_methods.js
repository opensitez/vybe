// vybe-test: js/classes/test_class_static_methods
// origin: languages/js/tests/js/js_classes_test.rs

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

class MathTools {
            static add(a, b) {
                return a + b;
            }
            static triple(v) {
                return MathTools.add(v, v + v);
            }
        }
        const total = MathTools.add(4, 5);
        const triple = MathTools.triple(7);
        __check(__line(total, triple), "9 21");
