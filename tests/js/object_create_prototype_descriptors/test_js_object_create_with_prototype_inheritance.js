// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_with_prototype_inheritance
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

const proto = { greeting: "Hello" };
const obj = Object.create(proto);
__check(__line(obj.greeting + "|isOwn=" + Object.hasOwn(obj, "greeting")), "Hello|isOwn=false");
