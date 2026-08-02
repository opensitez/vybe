// vybe-test: js/misc_advanced_patterns/reflect_vs_direct_access
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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

const obj = { x: 1 };
Object.defineProperty(obj, "y", { value: 2, configurable: false, writable: false, enumerable: true });
__check(__line(Reflect.get(obj, "x")), "1");
__check(__line(Reflect.has(obj, "y")), "true");
__check(__line(Reflect.ownKeys(obj).join(",")), "x,y");
__check(__line(Reflect.deleteProperty(obj, "x")), "true");
__check(__line(Reflect.deleteProperty(obj, "y")), "false");
