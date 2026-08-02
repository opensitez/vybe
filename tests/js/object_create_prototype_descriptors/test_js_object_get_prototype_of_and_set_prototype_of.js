// vybe-test: js/object_create_prototype_descriptors/test_js_object_get_prototype_of_and_set_prototype_of
// origin: languages/js/tests/js/test_js_object_create_prototype_descriptors.rs

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

const p1 = { a: 1 };
const p2 = { b: 2 };
const obj = Object.create(p1);
__check(__line(obj.a + "|b=" + obj.b), "1|b=undefined");
Object.setPrototypeOf(obj, p2);
__check(__line("a=" + obj.a + "|b=" + obj.b), "a=undefined|b=2");
