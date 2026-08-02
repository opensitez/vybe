// vybe-test: js/scope_closures_edge/const_object_mutable
// origin: languages/js/tests/js/test_scope_closures_edge.rs

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

const obj = { x: 1 };
obj.x = 99;
obj.y = 2;
__check(__line(obj.x), "99");
__check(__line(obj.y), "2");
