// vybe-test: js/proxy_traps_advanced/proxy_as_type_validator
// origin: languages/js/tests/js/test_proxy_traps_advanced.rs

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

function typed(obj, schema) {
    return new Proxy(obj, {
        set(target, prop, value) {
            if (schema[prop] && typeof value !== schema[prop]) {
                throw new TypeError(`${prop} must be ${schema[prop]}`);
            }
            target[prop] = value;
            return true;
        }
    });
}
const person = typed({}, { name: "string", age: "number" });
person.name = "Alice";
person.age = 30;
__check(__line(person.name + ":" + person.age), "Alice:30");
let threw = false;
try { person.age = "old"; } catch { threw = true; }
__check(__line(threw), "true");
