// vybe-test: js/coercion_modern/map_from_array
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

let m = new Map([["a", 1], ["b", 2], ["c", 3]]);
__check(__line(m.size), "3");
__check(__line(m.get("b")), "2");
