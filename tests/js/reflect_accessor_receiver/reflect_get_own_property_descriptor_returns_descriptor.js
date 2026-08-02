// vybe-test: js/reflect_accessor_receiver/reflect_get_own_property_descriptor_returns_descriptor
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

const o={}; Object.defineProperty(o,"x",{value:1,enumerable:true}); const d=Reflect.getOwnPropertyDescriptor(o,"x"); __check(__line(d.value), "1");__check(__line(d.enumerable), "true");
