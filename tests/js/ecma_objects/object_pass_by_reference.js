// vybe-test: js/ecma_objects/object_pass_by_reference
// origin: languages/js/tests/js/test_ecma_objects.rs

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

function modify(obj) {
    obj.x = 99;
}
const o = { x: 1 };
modify(o);
__check(__line(o.x), "99");
