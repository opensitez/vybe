// vybe-test: js/reflect_api/reflect_own_keys_all_types
// origin: languages/js/tests/js/test_reflect_api.rs

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
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Reflect.ownKeys(obj);
__check(__line(keys.includes("a")), "true");
__check(__line(keys.includes("hidden")), "true");
__check(__line(keys.some(k => typeof k === "symbol")), "true");
