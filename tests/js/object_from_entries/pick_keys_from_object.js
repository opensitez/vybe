// vybe-test: js/object_from_entries/pick_keys_from_object
// origin: languages/js/tests/js/test_object_from_entries.rs

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

function pick(obj, ...keys) {
    return Object.fromEntries(
        keys.filter(k => k in obj).map(k => [k, obj[k]])
    );
}
const user = { id: 1, name: "Alice", password: "secret", email: "alice@example.com" };
const safe = pick(user, "id", "name", "email");
__check(__line(Object.keys(safe).sort().join(",")), "email,id,name");
__check(__line(safe.name), "Alice");
__check(__line("password" in safe), "false");
