// vybe-test: js/scope_prototype/reflect_get_set
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let obj = { x: 1 };
__check(__line(Reflect.get(obj, "x")), "1");
Reflect.set(obj, "y", 2);
__check(__line(obj.y), "2");
