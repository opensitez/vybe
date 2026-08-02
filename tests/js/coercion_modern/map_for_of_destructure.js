// vybe-test: js/coercion_modern/map_for_of_destructure
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

let m = new Map([["x", 10], ["y", 20]]);
for (let [k, v] of m) {
    console.log(k + ":" + v);
}
