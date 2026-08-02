// vybe-test: js/features/test_switch
// origin: languages/js/tests/js/js_features_test.rs

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

let x = 2;
        switch (x) {
            case 1: console.log("one"); break;
            case 2: console.log("two"); break;
            case 3: console.log("three"); break;
            default: console.log("other");
        }
