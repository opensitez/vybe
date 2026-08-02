// vybe-test: js/scope_prototype/test_scope_prototype_reflect_ownkeys_symbol
// origin: languages/js/tests/js/test_scope_prototype.rs

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

const sym = Symbol("s");
const obj = { a: 1, [sym]: 2 };
const keys = Reflect.ownKeys(obj);
__check(__line(keys.length + "|" + (keys[1] === sym)), "2|true");
