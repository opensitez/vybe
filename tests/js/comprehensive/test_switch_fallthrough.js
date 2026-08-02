// vybe-test: js/comprehensive/test_switch_fallthrough
// origin: languages/js/tests/js/js_comprehensive_test.rs

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
        switch (2) {
            case 1: result = "one"; break;
            case 2: result = "two"; break;
            case 3: result = "three"; break;
        }
        __check(__line(result), "two");
