// vybe-test: js/weakmap_weakset_patterns/weakset_brand_checking
// origin: languages/js/tests/js/test_weakmap_weakset_patterns.rs

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

const validInstances = new WeakSet();
class Secure {
    constructor() {
        validInstances.add(this);
    }
    static validate(obj) {
        if (!validInstances.has(obj)) throw new TypeError("Invalid instance");
        return true;
    }
}
const s = new Secure();
__check(__line(Secure.validate(s)), "true");
let threw = false;
try { Secure.validate({}); } catch { threw = true; }
__check(__line(threw), "true");
