// vybe-test: js/class_patterns/class_extends_expression_is_evaluated_once
// origin: languages/js/tests/js/test_class_patterns.rs

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

let buildCount = 0;
function makeBase() {
    buildCount++;
    return class {
        greet() {
            return "from base";
        }
    };
}

class Child extends makeBase() {}
__check(__line(buildCount), "1");
__check(__line(new Child().greet()), "from base");
