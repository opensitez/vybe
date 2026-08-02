// vybe-test: js/reflect_accessor_receiver/reflect_set_integer_indexed_on_typed_array
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

const a=new Uint8Array(1); __check(__line(Reflect.set(a,0,9)), "true");__check(__line(a[0]), "9");
