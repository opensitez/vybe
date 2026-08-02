// vybe-test: js/reflect_api/reflect_has_checks_prototype_chain
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

const proto = { inherited: true };
const obj = Object.create(proto);
obj.own = true;
__check(__line(Reflect.has(obj, "own")), "true");
__check(__line(Reflect.has(obj, "inherited")), "true");
__check(__line(Reflect.has(obj, "missing")), "false");
