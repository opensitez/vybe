// vybe-test: js/ecma/test_for_of_with_break
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

let arr = [1, 2, 3, 4, 5];
        let sum = 0;
        for (let x of arr) {
            if (x > 3) break;
            sum = sum + x;
        }
        console.log(sum);
