// vybe-test: js/object_immutability/prevent_extensions_allows_modify
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

const obj = Object.preventExtensions({ x: 1 });
obj.x = 99;  // existing properties can be modified
obj.y = 2;   // new properties silently fail
__check(__line(obj.x), "99");
__check(__line(obj.y), "undefined");
