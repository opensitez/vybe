// vybe-test: js/immutable_patterns/immutable_update_object
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

const user = Object.freeze({ name: "Alice", age: 30, active: true });
const updated = { ...user, age: 31 };
__check(__line(user.age), "30");     // original unchanged
__check(__line(updated.age), "31");
__check(__line(updated.name), "Alice");
