// vybe-test: js/stdlib/test_combined_string_split_loop
// origin: languages/js/tests/js/js_stdlib_test.rs

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

let csv = "Alice,30,Bob,25,Charlie,35";
        let parts = csv.split(",");
        let names = [];
        for (let i = 0; i < parts.length; i += 2) {
            names.push(parts[i]);
        }
        console.log(names.join(" "));
