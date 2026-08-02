// vybe-test: js/host_interop/array_is_array
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

__check(__line(Array.isArray([1,2])), "true");
        __check(__line(Array.isArray("hello")), "false");
        __check(__line(Array.isArray(42)), "false");
