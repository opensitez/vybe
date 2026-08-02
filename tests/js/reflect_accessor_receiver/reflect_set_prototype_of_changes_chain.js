// vybe-test: js/reflect_accessor_receiver/reflect_set_prototype_of_changes_chain
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

const o={}; const p={tag:"p"}; Reflect.setPrototypeOf(o,p); __check(__line(Reflect.getPrototypeOf(o)===p), "true");
