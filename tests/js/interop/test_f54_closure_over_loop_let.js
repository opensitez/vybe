// vybe-test: js/interop/test_f54_closure_over_loop_let
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

let fns = [];
        for (let i = 0; i < 5; i++) {
            fns.push(() => i);
        }
        console.log(fns[0](), fns[1](), fns[2](), fns[3](), fns[4]());
