// vybe-test: js/object_advanced/object_keys_no_inherited
// origin: languages/js/tests/js/test_object_advanced.rs

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

let parent = { x: 1 };
let child = Object.create(parent);
child.y = 2;
child.z = 3;
__check(__line(Object.keys(child).join(",")), "y,z");
