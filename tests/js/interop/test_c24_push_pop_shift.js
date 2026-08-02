// vybe-test: js/interop/test_c24_push_pop_shift
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

let arr = [1, 2, 3];
        arr.push(4);
        let popped = arr.pop();
        let shifted = arr.shift();
        __check(__line(popped, shifted, arr), "4 1 2,3");
