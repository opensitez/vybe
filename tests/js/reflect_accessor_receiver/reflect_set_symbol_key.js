// vybe-test: js/reflect_accessor_receiver/reflect_set_symbol_key
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

const s=Symbol("k"); const o={}; __check(__line(Reflect.set(o,s,8)), "true");__check(__line(o[s]), "8");
