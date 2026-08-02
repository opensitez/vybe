// vybe-test: js/reflect_accessor_receiver/reflect_prevent_extensions_blocks_add
// origin: languages/js/tests/js/test_reflect_accessor_receiver.rs

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

const o={x:1}; Reflect.preventExtensions(o); __check(__line(Reflect.set(o,"y",2)), "false");__check(__line("y" in o), "false");
