// vybe-test: js/reflect_accessor_receiver/reflect_define_property_on_extensible_object_adds_symbol
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

const s=Symbol("d"); const o={}; Reflect.defineProperty(o,s,{value:1}); __check(__line(Reflect.get(o,s)), "1");
