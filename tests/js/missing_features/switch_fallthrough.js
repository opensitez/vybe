// vybe-test: js/missing_features/switch_fallthrough
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

let x = 1;
        let result = "";
        switch (x) {
            case 1:
                result += "one";
            case 2:
                result += "two";
            case 3:
                result += "three";
        }
        __check(__line(result), "onetwothree");
