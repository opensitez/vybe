// vybe-test: js/advanced/test_nested_if
// origin: languages/js/tests/js/js_advanced_test.rs

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

function classify(n) {
            if (n > 0) {
                if (n > 100) return "big";
                else return "small positive";
            } else if (n < 0) {
                return "negative";
            } else {
                return "zero";
            }
        }
        __check(__line(classify(50), classify(-5), classify(0), classify(200)), "small positive negative zero big");
