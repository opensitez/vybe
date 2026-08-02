// vybe-test: js/coercion_modern/coerce_truthy_values
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let truthies = [true, 1, -1, "hello", {}, [], "false", "0"];
truthies.forEach(v => console.log(Boolean(v)));
