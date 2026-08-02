// vybe-test: js/delete_operator/delete_inherited_has_no_effect
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

const proto = { x: 1 };
const obj = Object.create(proto);
delete obj.x; // x is on proto, not own
__check(__line("x" in obj), "true"); // still accessible via prototype
__check(__line(obj.hasOwnProperty("x")), "false");
