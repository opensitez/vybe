// vybe-test: js/ecma_iterators/for_loop_c_style
// origin: languages/js/tests/js/test_ecma_iterators.rs

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

let sum = 0;
for (let i = 0; i < 5; i++) {
    sum += i;
}
console.log(sum);
