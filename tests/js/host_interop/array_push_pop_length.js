// vybe-test: js/host_interop/array_push_pop_length
// origin: languages/js/tests/js/js_host_interop_test.rs

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

let a = [1, 2];
        a.push(3);
        __check(__line(a.length), "3");
        let x = a.pop();
        __check(__line(x), "3");
        __check(__line(a.length), "2");
