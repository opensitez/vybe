// vybe-test: js/interop/test_b20_delete_property
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

let obj = { a: 1, b: 2, c: 3 };
        delete obj.b;
        __check(__line("b" in obj, obj.a, obj.c), "false 1 3");
