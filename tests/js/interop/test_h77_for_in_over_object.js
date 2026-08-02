// vybe-test: js/interop/test_h77_for_in_over_object
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

let obj = { a: 1, b: 2, c: 3 };
        let sum = 0;
        for (let k in obj) {
            sum += obj[k];
        }
        console.log(sum);
