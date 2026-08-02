// vybe-test: js/object_advanced/object_assign_override_order
// origin: languages/js/tests/js/test_object_advanced.rs

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

let defaults = { color: "red", size: 10, bold: false };
let user = { color: "blue", bold: true };
let result = Object.assign({}, defaults, user);
__check(__line(result.color), "blue");
__check(__line(result.size), "10");
__check(__line(result.bold), "true");
