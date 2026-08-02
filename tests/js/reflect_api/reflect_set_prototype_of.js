// vybe-test: js/reflect_api/reflect_set_prototype_of
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

const a = { hello() { return "a"; } };
const b = { hello() { return "b"; } };
const obj = Object.create(a);
__check(__line(obj.hello()), "a");
Reflect.setPrototypeOf(obj, b);
__check(__line(obj.hello()), "b");
