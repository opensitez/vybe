// vybe-test: js/closures_functional/closure_over_let_in_loop
// origin: languages/js/tests/js/test_closures_functional.rs

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

let funcs = [];
for (let i = 0; i < 5; i++) {
    funcs.push(() => i);
}
console.log(funcs[0]());
console.log(funcs[2]());
console.log(funcs[4]());
