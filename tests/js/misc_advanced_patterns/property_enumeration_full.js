// vybe-test: js/misc_advanced_patterns/property_enumeration_full
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

const sym = Symbol("s");
const obj = Object.create({ inherited: true });
obj.own = 1;
Object.defineProperty(obj, "nonEnum", { value: 2, enumerable: false });
obj[sym] = "symbol";
__check(__line(Object.keys(obj).join(",")), "own");
__check(__line(Object.getOwnPropertyNames(obj).sort().join(",")), "nonEnum,own");
__check(__line(Reflect.ownKeys(obj).filter(k => typeof k === "string").sort().join(",")), "nonEnum,own");
