// vybe-test: js/ecma/test_split_and_for_of
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

let csv = "a,b,c,d";
        let parts = csv.split(",");
        let result = "";
        for (let p of parts) {
            result = result + p.toUpperCase() + " ";
        }
        console.log(result.trim());
