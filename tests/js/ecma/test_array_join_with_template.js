// vybe-test: js/ecma/test_array_join_with_template
// origin: languages/js/tests/js/js_ecma_test.rs

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

let names = ["Alice", "Bob", "Charlie"];
        __check(__line(`Names: ${names.join(", ")}`), "Names: Alice, Bob, Charlie");
