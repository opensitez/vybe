// vybe-test: js/object_immutability/seal_allows_modify_prevents_add_delete
// origin: languages/js/tests/js/test_object_immutability.rs

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

const obj = Object.seal({ x: 1, y: 2 });
obj.x = 99; // modify ok
obj.z = 3;  // add — silently fails
delete obj.y; // delete — silently fails
__check(__line(obj.x), "99");
__check(__line(obj.z), "undefined");
__check(__line(obj.y), "2");
