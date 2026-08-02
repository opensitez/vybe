// vybe-test: js/comprehensive/test_switch_with_break
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

let x = "b";
        switch (x) {
            case "a": console.log("A"); break;
            case "b": console.log("B"); break;
            case "c": console.log("C"); break;
        }
