// vybe-test: js/host_interop/array_sort_reverse_concat
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

let a = [3, 1, 2];
        a.sort((a, b) => a - b);
        __check(__line(a.join(",")), "1,2,3");
        a.reverse();
        __check(__line(a.join(",")), "3,2,1");
        let b = a.concat([4, 5]);
        __check(__line(b.join(",")), "3,2,1,4,5");
