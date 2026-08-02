// vybe-test: js/object_create_prototype_descriptors/test_js_object_is_prototype_of_chain_check
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

const grandParent = {};
const parent = Object.create(grandParent);
const child = Object.create(parent);

__check(__line(`${grandParent.isPrototypeOf(child)}:${parent.isPrototypeOf(child)}:${child.isPrototypeOf(parent)}`), "true:true:false");
