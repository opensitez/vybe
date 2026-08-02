// vybe-test: js/delete_operator/delete_computed_property
// origin: languages/js/tests/js/test_delete_operator.rs

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

const obj = { x: 1, y: 2 };
const key = "x";
delete obj[key];
__check(__line("x" in obj), "false");
__check(__line("y" in obj), "true");
