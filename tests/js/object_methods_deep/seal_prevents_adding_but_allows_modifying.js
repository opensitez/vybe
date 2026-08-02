// vybe-test: js/object_methods_deep/seal_prevents_adding_but_allows_modifying
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const obj = Object.seal({ x: 1 });
obj.y = 2; // adding fails
obj.x = 99; // modifying works
__check(__line("y" in obj), "false");
__check(__line(obj.x), "99");
