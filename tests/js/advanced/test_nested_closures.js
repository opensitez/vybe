// vybe-test: js/advanced/test_nested_closures
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

function outer() {
            let x = 10;
            function middle() {
                let y = 20;
                function inner() {
                    return x + y;
                }
                return inner();
            }
            return middle();
        }
        __check(__line(outer()), "30");
