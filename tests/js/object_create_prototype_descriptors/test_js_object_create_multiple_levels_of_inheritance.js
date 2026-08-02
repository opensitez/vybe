// vybe-test: js/object_create_prototype_descriptors/test_js_object_create_multiple_levels_of_inheritance
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

const l1 = { level: 1 };
const l2 = Object.create(l1);
l2.level = 2;
const l3 = Object.create(l2);
__check(__line(`${l3.level}:${l2.level}:${l1.level}`), "2:2:1");
