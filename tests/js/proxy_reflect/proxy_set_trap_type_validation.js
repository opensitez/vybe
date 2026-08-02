// vybe-test: js/proxy_reflect/proxy_set_trap_type_validation
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

const handler = {
    set(target, prop, value) {
        if (prop === "age" && (typeof value !== "number" || value < 0)) {
            throw new RangeError("age must be non-negative number");
        }
        target[prop] = value;
        return true;
    }
};
const person = new Proxy({}, handler);
person.age = 25;
__check(__line(person.age), "25");
let err = "";
try { person.age = -1; } catch(e) { err = e.message; }
__check(__line(err), "age must be non-negative number");
