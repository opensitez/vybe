// vybe-test: js/spread_rest_advanced/object_spread_add_new_properties
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

const user = { name: "Alice" };
const withRole = { ...user, role: "admin" };
__check(__line(withRole.name), "Alice");
__check(__line(withRole.role), "admin");
