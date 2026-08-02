// vybe-test: js/interop/test_h73_switch_fallthrough
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

let result = "";
        switch (1) {
            case 1: result += "a";
            case 2: result += "b";
            case 3: result += "c";
        }
        __check(__line(result), "abc");
