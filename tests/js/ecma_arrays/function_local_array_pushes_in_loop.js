// vybe-test: js/ecma_arrays/function_local_array_pushes_in_loop
// origin: languages/js/tests/js/test_ecma_arrays.rs

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

function main() {
    const arr = [];
    for (let i = 1; i <= 3; i++) {
        arr.push(i);
    }
    console.log(arr.join(","));
}
main();
