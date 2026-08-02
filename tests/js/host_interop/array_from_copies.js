// vybe-test: js/host_interop/array_from_copies
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

let orig = [1, 2, 3];
        let copy = Array.from(orig);
        copy.push(4);
        __check(__line(orig.length), "3");
        __check(__line(copy.length), "4");
