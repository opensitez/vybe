// vybe-test: js/missing_features/switch_no_match_no_default
// origin: languages/js/tests/js/js_missing_features_test.rs

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

let result = "none";
        switch (99) {
            case 1: result = "one"; break;
            case 2: result = "two"; break;
        }
        __check(__line(result), "none");
