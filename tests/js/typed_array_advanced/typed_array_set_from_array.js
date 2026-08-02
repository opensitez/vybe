// vybe-test: js/typed_array_advanced/typed_array_set_from_array
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const ta = new Int32Array(5);
ta.set([10, 20, 30]);
__check(__line(ta[0]), "10");
__check(__line(ta[1]), "20");
__check(__line(ta[2]), "30");
__check(__line(ta[3]), "0"); // unfilled — 0
