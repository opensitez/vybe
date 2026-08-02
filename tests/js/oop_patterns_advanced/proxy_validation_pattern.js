// vybe-test: js/oop_patterns_advanced/proxy_validation_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

function createValidated(target, validators) {
    return new Proxy(target, {
        set(obj, prop, value) {
            if (validators[prop] && !validators[prop](value)) throw new Error(`Invalid ${prop}`);
            obj[prop] = value;
            return true;
        }
    });
}
const person = createValidated({}, { age: v => typeof v === "number" && v >= 0 && v <= 150 });
person.name = "Alice";
person.age = 30;
__check(__line(person.name), "Alice");
__check(__line(person.age), "30");
let threw = false;
try { person.age = -5; } catch { threw = true; }
__check(__line(threw), "true");
