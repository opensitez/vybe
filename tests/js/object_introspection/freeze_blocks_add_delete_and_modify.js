// vybe-test: js/object_introspection/freeze_blocks_add_delete_and_modify
// origin: languages/js/tests/js/test_object_introspection.rs

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

const obj = { a: 1, b: 2 };
Object.freeze(obj);
obj.a = 99;
obj.c = 3;
delete obj.b;
__check(__line(obj.a), "1");
__check(__line(obj.b), "2");
__check(__line(obj.c), "undefined");
