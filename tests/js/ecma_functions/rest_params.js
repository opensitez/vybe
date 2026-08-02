// vybe-test: js/ecma_functions/rest_params
// origin: languages/js/tests/js/test_ecma_functions.rs

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

function sum(...nums) {
    let total = 0;
    for (const n of nums) {
        total += n;
    }
    return total;
}
console.log(sum(1, 2, 3, 4));
