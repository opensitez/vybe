// vybe-test: js/comprehensive/test_json_roundtrip
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

let original = [1, 2, 3];
        let copy = JSON.parse(JSON.stringify(original));
        __check(__line(copy), "1,2,3");
