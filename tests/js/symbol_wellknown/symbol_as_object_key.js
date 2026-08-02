// vybe-test: js/symbol_wellknown/symbol_as_object_key
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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

const id = Symbol("id");
const user = { [id]: 42, name: "Alice" };
__check(__line(user[id]), "42");
__check(__line(user.name), "Alice");
