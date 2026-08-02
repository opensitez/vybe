// vybe-test: js/ecma/test_default_param_expression
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

function makeArray(size = 3, fill = 0) {
            let arr = [];
            for (let i = 0; i < size; i++) { arr.push(fill); }
            return arr;
        }
        console.log(makeArray());
        console.log(makeArray(2, 7));
