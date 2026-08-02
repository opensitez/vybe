// vybe-test: js/reflect_accessor_receiver/reflect_own_keys_on_array_includes_length
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

const k=Reflect.ownKeys([1,2]); __check(__line(k.includes("length")), "true");__check(__line(k.includes("0")), "true");
