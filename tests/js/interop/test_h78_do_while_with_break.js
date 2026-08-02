// vybe-test: js/interop/test_h78_do_while_with_break
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

let i = 0;
        let sum = 0;
        do {
            sum += i;
            i++;
            if (i > 5) break;
        } while (i < 100);
        console.log(sum);
