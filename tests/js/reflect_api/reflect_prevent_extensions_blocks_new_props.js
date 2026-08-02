// vybe-test: js/reflect_api/reflect_prevent_extensions_blocks_new_props
// origin: languages/js/tests/js/test_reflect_api.rs

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
Reflect.preventExtensions(obj);
obj.y = 2; // silently fails
__check(__line("y" in obj), "false");
obj.x = 99; // existing props still modifiable
__check(__line(obj.x), "99");
