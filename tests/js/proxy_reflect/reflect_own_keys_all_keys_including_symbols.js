// vybe-test: js/proxy_reflect/reflect_own_keys_all_keys_including_symbols
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

const sym = Symbol("id");
const obj = { a: 1, b: 2 };
obj[sym] = 99;
const keys = Reflect.ownKeys(obj);
// All 3 keys present: "a", "b", and the symbol
__check(__line(keys.length), "3");
__check(__line(keys.includes("a")), "true");
__check(__line(keys.includes("b")), "true");
