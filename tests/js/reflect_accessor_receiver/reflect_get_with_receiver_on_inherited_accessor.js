// vybe-test: js/reflect_accessor_receiver/reflect_get_with_receiver_on_inherited_accessor
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

const base={_v:1,get g(){return this._v;}}; const o=Object.create(base); __check(__line(Reflect.get(o,"g",{_v:100})), "100");
