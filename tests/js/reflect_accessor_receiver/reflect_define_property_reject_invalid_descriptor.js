// vybe-test: js/reflect_accessor_receiver/reflect_define_property_reject_invalid_descriptor
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

const o={}; Object.defineProperty(o,"x",{value:1,configurable:false}); __check(__line(Reflect.defineProperty(o,"x",{value:2,configurable:true})), "false");
